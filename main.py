import uuid, asyncio, tempfile
from pathlib import Path
from urllib.parse import quote
from typing import Optional

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

from processing.data_cleaning import run_data_cleaning, extract_stats, compute_yoy
from processing.pptx_builder import generate_pptx

app = FastAPI(title="승강설비 장애분석 API v21")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

WORK_DIR  = Path(tempfile.gettempdir()) / "seungang_jobs"
WORK_DIR.mkdir(parents=True, exist_ok=True)
jobs: dict = {}
STATIC_DIR = Path(__file__).parent / "static"


@app.get("/health")
def health():
    return {"status": "ok", "version": "21"}


@app.post("/api/upload")
async def upload_files(
    background_tasks: BackgroundTasks,
    fault_report: UploadFile = File(...),
    fault_class:  UploadFile = File(...),
    fault_report_last: Optional[UploadFile] = File(None),   # ← 신규: 작년 파일 (선택)
):
    job_id  = str(uuid.uuid4())
    job_dir = WORK_DIR / job_id
    job_dir.mkdir(parents=True)

    src_path = job_dir / "fault_input.xlsx"
    cls_path = job_dir / "fault_class.xlsx"

    with open(src_path, "wb") as f: f.write(await fault_report.read())
    with open(cls_path, "wb") as f: f.write(await fault_class.read())

    # 작년 파일 (있으면 저장)
    last_path = None
    has_last  = False
    if fault_report_last and fault_report_last.filename:
        last_path = job_dir / "fault_input_last.xlsx"
        content   = await fault_report_last.read()
        if len(content) > 100:
            with open(last_path, "wb") as f: f.write(content)
            has_last = True

    jobs[job_id] = {"status": "processing", "step": "파일 업로드 완료", "progress": 5,
                    "has_yoy": has_last}
    background_tasks.add_task(_run_pipeline, job_id, job_dir, src_path, cls_path, last_path)
    return {"job_id": job_id}


@app.get("/api/status/{job_id}")
def get_status(job_id: str):
    if job_id not in jobs: raise HTTPException(404, "Job not found")
    return jobs[job_id]


@app.get("/api/download/{job_id}/{filename}")
def download_file(job_id: str, filename: str):
    try: uuid.UUID(job_id)
    except ValueError: raise HTTPException(400, "Invalid job_id")

    filename  = Path(filename).name
    file_path = WORK_DIR / job_id / filename
    if not file_path.exists(): raise HTTPException(404, f"파일을 찾을 수 없습니다: {filename}")

    saveas   = filename
    job_files = jobs.get(job_id, {}).get("files", [])
    for f in job_files:
        if f.get("filename") == filename:
            saveas = f.get("saveas", filename); break

    encoded_name = quote(saveas, safe="()-_.")
    headers = {"Content-Disposition": "attachment; filename*=UTF-8''" + encoded_name}
    return FileResponse(path=str(file_path), media_type="application/octet-stream", headers=headers)


@app.get("/", include_in_schema=False)
def serve_index():
    idx = STATIC_DIR / "index.html"
    if idx.exists(): return FileResponse(str(idx), media_type="text/html")
    return JSONResponse({"message": "승강설비 장애분석 API 실행 중"})

@app.get("/{full_path:path}", include_in_schema=False)
def serve_static(full_path: str):
    idx = STATIC_DIR / "index.html"
    if idx.exists(): return FileResponse(str(idx), media_type="text/html")
    return JSONResponse({"message": "Static files not found."})


async def _run_pipeline(job_id: str, job_dir: Path,
                         src_path: Path, cls_path: Path, last_path):
    loop = asyncio.get_event_loop()
    try:
        out_xlsx = job_dir / "fault_report.xlsx"
        el_pptx  = job_dir / "el_report.pptx"
        es_pptx  = job_dir / "es_report.pptx"

        jobs[job_id].update({"step": "데이터 정제 중...", "progress": 12})
        await loop.run_in_executor(None, run_data_cleaning, src_path, cls_path, out_xlsx)
        if not out_xlsx.exists(): raise RuntimeError("데이터 정제 실패")

        jobs[job_id].update({"step": "데이터 분석 중...", "progress": 30})
        stats = await loop.run_in_executor(None, extract_stats, out_xlsx)

        # ── 전년 동기 비교 (작년 파일이 있을 때만)
        yoy = None
        if last_path and last_path.exists():
            jobs[job_id].update({"step": "전년 동기 비교 분석 중...", "progress": 42})
            last_xlsx = job_dir / "fault_report_last.xlsx"
            await loop.run_in_executor(None, run_data_cleaning, last_path, cls_path, last_xlsx)
            if last_xlsx.exists():
                stats_last = await loop.run_in_executor(None, extract_stats, last_xlsx)
                yoy = compute_yoy(stats, stats_last)

        jobs[job_id].update({"step": "엘리베이터 보고서 생성 중...", "progress": 55})
        await loop.run_in_executor(None, generate_pptx, stats, "EL", str(el_pptx), yoy)

        jobs[job_id].update({"step": "에스컬레이터 보고서 생성 중...", "progress": 78})
        await loop.run_in_executor(None, generate_pptx, stats, "ES", str(es_pptx), yoy)

        year    = stats.get("year",    2025)
        quarter = stats.get("quarter", "4분기")
        suffix  = f"({year}년_{quarter})"
        dynamic_files = [
            {"filename": "fault_report.xlsx", "label": "정제 데이터 (xlsx)",
             "saveas": f"장애신고_처리결과{suffix}.xlsx"},
            {"filename": "el_report.pptx", "label": "EL 장애분석 보고서",
             "saveas": f"엘리베이터_장애분석보고서{suffix}.pptx"},
            {"filename": "es_report.pptx", "label": "ES 장애분석 보고서",
             "saveas": f"에스컬레이터_장애분석보고서{suffix}.pptx"},
        ]

        jobs[job_id].update({
            "status": "done", "step": "완료", "progress": 100,
            "files":  [{"name": f["filename"], "label": f["label"], "saveas": f["saveas"]}
                       for f in dynamic_files],
            "stats":  stats,
            "has_yoy": yoy is not None,
        })

    except Exception as e:
        jobs[job_id].update({"status": "error", "step": f"오류: {e}", "progress": 0})
