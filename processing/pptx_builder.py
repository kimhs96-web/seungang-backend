"""
PPT 보고서 생성 모듈

핵심 설계 원칙:
  - JS 파일을 backend/ 폴더에 고정 생성하고 backend/ 에서 실행
  - backend/node_modules/pptxgenjs 가 있으면 무조건 로컬로 찾음
  - 경로를 JS require()에 하드코딩하지 않음
  - Windows/Mac/Linux 모두 동일하게 동작
"""

import json
import math
import subprocess
import sys
import os
from pathlib import Path

# backend 폴더 = 이 파일(processing/pptx_builder.py)의 부모의 부모
BACKEND_DIR = Path(__file__).resolve().parent.parent


def _run_js(js_code: str, label: str) -> None:
    """
    JS 코드를 backend/ 폴더에 파일로 저장하고 node로 실행.
    backend/ 에서 실행하므로 node_modules/pptxgenjs 를 항상 찾음.
    """
    script_path = BACKEND_DIR / f"_runner_{label}.js"
    script_path.write_text(js_code, encoding='utf-8')

    # node 실행 파일 경로
    import shutil
    node_exe = shutil.which("node") or "node"

    # Windows: shell=True 필요 (node.cmd 인식)
    use_shell = (sys.platform == "win32")

    if use_shell:
        cmd = f'"{node_exe}" "{script_path}"'
    else:
        cmd = [node_exe, str(script_path)]

    result = subprocess.run(
        cmd,
        cwd=str(BACKEND_DIR),   # ← 핵심: backend/ 에서 실행
        capture_output=True,
        text=True,
        timeout=180,
        shell=use_shell,
        encoding='utf-8',
        errors='replace',
    )

    # 임시 JS 파일 삭제
    try:
        script_path.unlink()
    except Exception:
        pass

    if result.returncode != 0:
        raise RuntimeError(result.stderr[:1000])


# ── 장치별 원인·대책 ──────────────────────────────────────────────────────────

EL_CAUSES = {
    '안전스위치':   ('광센서·포토센서 라인 불량, 카도어 스위치 간격 이상, 칼리브레이션 오류',
                    '에러소거 후 재가동, 스위치 간격 조정, 라인 재연결',
                    '정기점검 시 안전스위치 간격·접점 상태 확인, 광센서 라인 상태 주기적 점검, 예비품 확보',
                    'C0392B'),
    '전장품':       ('전자접촉기·마그네트 접점 불량, 파워서플라이·SMPS 불량',
                    '접점 보수·교체, 마그네트 교체, 파워서플라이 교체',
                    '접점류 정기 확인 및 청소, 전장품 이상 시 즉시 교체, 주요 전장품 예비품 재고 유지',
                    'E8843A'),
    '제어장치':     ('메인기판(PCB) 불량, 인버터 오류, 데이터 이상',
                    '기판 교체, 데이터 재입력·에러소거, 인버터 점검',
                    '기판류 데이터 백업 및 정기 점검, 인버터 환기팬 작동 확인, 결선 접촉 불량 예방 관리',
                    '2C5F8A'),
    '도어':         ('행거롤러 파손, 도어 기판 에러, 세이프티슈 이상, 가이드롤러 소음',
                    '롤러 교체, 도어기판 교체, 칼리브레이션 실시, 간격 조정',
                    '도어 롤러·레일 정기 점검·주유, 도어 기판 노후도 확인 후 예방 교체',
                    '0D7A8A'),
    '제동장치':     ('브레이크 플런저 개방 불량, 브레이크 소음·간격 이상',
                    '플런저 간격 조정, 브레이크 간격 조정 후 재가동',
                    '점검 시 브레이크 동작 간격 정기 확인, 소음·진동 이상 감지 즉시 조치',
                    'F5A623'),
    '속도조정장치': ('인버터 에러 코드 발생(엔코더 펄스 이상 등), 오토튜닝 미흡',
                    '인버터 에러소거·오토튜닝, 엔코더 라인 전압 조정',
                    '인버터 설정값·단자 라인 정기 확인, 환기팬 점검, 노후 시 예방 교체',
                    '27AE60'),
    '층감지장치':   ('층감지 센서 이상, 착상장치 오류',
                    '센서 점검·교체, 착상장치 조정',
                    '센서류 부품 상태 확인·조정, 적기 교체를 통한 예방 관리',
                    '8E44AD'),
    '동력장치':     ('모터·감속기 이상, 권상기 베어링 마모',
                    '베어링 교체, 모터 점검',
                    '베어링 등 부품 이상여부 확인, 소음·진동 점검을 통한 적기 예방 교체',
                    '16A085'),
    '동력전달장치': ('시브·로프·베어링 이상',
                    '시브베어링 점검, 체인·로프 교체',
                    '시브베어링·체인 등 장치류 소음·작동상태 확인, 예비품 확보로 신속 보수',
                    '2980B9'),
    '부속장치':     ('인터폰 자기 불량, 단자대 접촉 이상, 층표시 불량',
                    '자기 교체, 단자대 재취부, 부속 보수',
                    '부속장치 정기 점검 및 상태 확인, 소모성 부품 예비품 확보',
                    '64748B'),
}

ES_CAUSES = {
    '안전스위치':   ('스텝처짐 스위치 간격 이상, 브레이크 개방확인 스위치 간격 미흡, 역회전 감지 센서 오작동',
                    '스위치 간격 조정, 에러소거 후 재가동, a·b상 센서 상태 확인',
                    '정기점검 시 안전스위치 전수 간격 측정·조정, 브레이크 개방 스위치 간격 기준값 관리 강화',
                    'C0392B'),
    '제어장치':     ('PCB·PLC 기판 불량, 역주행 감지기판 에러, 통신 불량',
                    '기판 교체, 에러소거·학습운전 실시, 통신 라인 점검',
                    '제어기판 이상이력 관리 및 데이터 백업, PLC 파라미터 정기 확인',
                    '2C5F8A'),
    '핸드레일':     ('핸드레일 장력 편차, 곡각부 베어링 파손, 텐션롤러 마모',
                    '핸드레일 장력 조정, 텐션 베어링 교체, 곡각부 청수·교체',
                    '핸드레일 장력 정기 측정 및 기준값 유지, 곡각부 베어링 수명 관리·예방 교체',
                    '0D8A6A'),
    '속도조정장치': ('인버터 에러(설정값 이상, 과열, 파라미터 오류), 인버터 파손',
                    '인버터 에러소거, 가속·감속 설정값 변경, 인버터 교체',
                    '인버터 설정값 정기 점검 및 기록, 방열 환경 관리, 노후 인버터 예방 교체 계획',
                    'E8843A'),
    '제동장치':     ('보조브레이크 작동, 브레이크 개방 불량, 라이닝 마모',
                    '보조브레이크 복귀 및 안전점검, 브레이크 갭 조정',
                    '브레이크 개방 간격 정기 측정·조정, 라이닝 마모도 확인 및 예방 교체',
                    'F5A623'),
    '동력전달장치': ('곡각부 마모, 가압롤러 파손, 스텝체인 롤러 파손, 게이트롤러 베어링 불량',
                    '곡각부 교체, 가압롤러 교체, 스텝체인 롤러 교체',
                    '스텝체인·롤러류 윤활 및 마모 상태 정기 점검, 예비품 재고 확보로 신속 보수',
                    '1B6B4A'),
    '스텝콤':       ('스커트 가드 변형, 스텝-스커트 간섭·마찰 소음, 가이드레일 이상',
                    '스커트 가드 간격 조정·고정, 윤활유 주입, 스텝 치수 확인',
                    '스텝-스커트 간격 기준 유지, 소음 발생 시 즉시 윤활·조정',
                    '6B4C9A'),
    '부속장치':     ('스커트 간섭 소음, 배수펌프 불량, 음성안내장치 이상',
                    '스커트 간격 조정, 배수펌프 교체, 부속 보수',
                    '부속장치류 정기 점검 및 상태 확인, 소모성 부품 예비품 확보',
                    '4A5568'),
    '동력장치':     ('모터 베어링 마모·불량, 모터 과열',
                    '모터 베어링 교체(추후), 냉각 상태 점검',
                    '모터 베어링 수명 예측 관리 및 예방 교체',
                    '2AAF7E'),
    '전장품':       ('전자접촉기 작동 불량, 마그네트 접점 이상, 누전차단기 트립',
                    '전자접촉기 교체, 접점 확인·보수, 누전차단기 복귀',
                    '전장품 접점 정기 확인·청소, 전장함 방수 상태 확인',
                    '7BC67E'),
    '멀티포스트':   ('포토센서 불량, 안내장치 이상',
                    '포토센서 교체, 안내장치 점검',
                    '포토센서 정기 점검 및 청소, 안내장치 상태 확인',
                    '95A5A6'),
}


# ── 공개 API ──────────────────────────────────────────────────────────────────

def generate_pptx(stats: dict, report_type: str, output_path: str, yoy: dict = None) -> None:
    causes = EL_CAUSES if report_type == 'EL' else ES_CAUSES
    data   = stats['el'] if report_type == 'EL' else stats['es']

    l1, l2 = data['line1'], data['line2']
    total   = l1['total'] + l2['total']
    avg_rec = round(
        (l1['avgRecovery'] * l1['total'] + l2['avgRecovery'] * l2['total']) / max(total, 1), 1
    )

    all_faults = set(list(l1['faults'].keys()) + list(l2['faults'].keys()))
    fault_list = sorted(
        [(k, causes[k]) for k in causes if k in all_faults],
        key=lambda x: l1['faults'].get(x[0], 0) + l2['faults'].get(x[0], 0),
        reverse=True,
    )

    js_code = _build_js(stats, report_type, l1, l2, total, avg_rec, fault_list, output_path, yoy)
    label   = report_type.lower()

    try:
        _run_js(js_code, label)
    except RuntimeError as e:
        raise RuntimeError(f"{report_type} PPT 생성 실패:\n{e}")


# ── JS 생성 헬퍼 ──────────────────────────────────────────────────────────────

def _q(s: str) -> str:
    """Python 문자열 → JS 작은따옴표 문자열 (안전 이스케이프)"""
    return "'" + str(s).replace('\\', '\\\\').replace("'", "\\'").replace('\n', '\\n') + "'"


def _build_analysis_opinion(quarter, l1_total, l2_total, monthly_trend,
                             top_fault, l2_top, l1_faults, l2_faults,
                             l1_months, l1_vals):
    """배포 전 보고서와 동일한 상세 분석 의견 텍스트 생성
    실제 줄바꿈(\n)을 사용하면 _q()가 \\n으로 변환하여 JS에서 올바르게 처리됨
    """
    # 1호선 월별 추이 방향 파악
    if len(l1_vals) >= 2:
        if l1_vals[-1] > l1_vals[0]:
            trend_dir = "(증가 추세)"
        elif l1_vals[-1] < l1_vals[0]:
            trend_dir = "(감소 추세)"
        else:
            trend_dir = "(유지)"
    else:
        trend_dir = ""

    # 1호선 상위 2개 장애
    l1_top2 = list(l1_faults.keys())[:2]
    l1_top2_str = "·".join(l1_top2) + " 장애 집중" if l1_top2 else top_fault

    # 2호선 1위 장애 비율
    l2_total_safe = max(l2_total, 1)
    l2_top_cnt  = l2_faults.get(l2_top, 0)
    l2_top_pct  = round(l2_top_cnt / l2_total_safe * 100, 1)

    # 2호선 2위 장애
    l2_keys = list(l2_faults.keys())
    l2_2nd  = l2_keys[1] if len(l2_keys) > 1 else ""
    l2_2nd_str = f"  {l2_2nd} 장애 2위로 집중 관리 필요" if l2_2nd else ""

    # 최다 발생 월
    if l1_months and l1_vals:
        max_idx    = l1_vals.index(max(l1_vals))
        peak_month = l1_months[max_idx]
    else:
        peak_month = quarter[:2]

    # 실제 \n 줄바꿈 사용 (1호선/2호선 사이 빈 줄 포함)
    text = (
        f"▶ 1호선 : {quarter} 총 {l1_total}건 발생 {trend_dir}\n"
        f"  월별 추이 : {monthly_trend}\n"
        f"  {l1_top2_str}\n"
        f"\n"
        f"▶ 2호선 : {quarter} 총 {l2_total}건 발생\n"
        f"  {l2_top} 장애가 전체의 {l2_top_pct}% 차지"
    )
    if l2_2nd_str:
        text += f"\n{l2_2nd_str}"
    text += (
        f"\n\n"
        f"⇒ {peak_month} 증가 원인 : 동절기 온도변화에 의한\n"
        f"  전장품·센서류 오동작 증가 추정"
    )

    return _q(text)


def _j(v) -> str:
    """Python 값 → JSON 리터럴"""
    return json.dumps(v, ensure_ascii=False)


def _pct(a, b) -> float:
    return round(a / max(b, 1) * 100, 1)


def _build_js(stats, rt, l1, l2, total, avg_rec, fault_list, output_path, yoy=None) -> str:
    is_el   = (rt == 'EL')
    quarter = stats['quarter']
    year    = stats['year']
    period  = stats['period']

    # 색상
    if is_el:
        pri, acc, sec, ter = '1B2A4A', 'E8843A', '2C5F8A', '0D7A8A'
        cover_bg, top_bar  = '1B2A4A', '2C5F8A'
        slide_bg, foot_c   = 'F7F9FB', 'AABBD0'
    else:
        pri, acc, sec, ter = '1A3C34', '2AAF7E', '1B6B4A', '0D8A6A'
        cover_bg, top_bar  = '1A3C34', '1B6B4A'
        slide_bg, foot_c   = 'F5FAF7', '88BBAA'

    title_kr = '엘리베이터' if is_el else '에스컬레이터'
    title_en = 'Elevator' if is_el else 'Escalator'
    label_cd = 'EL' if is_el else 'ES'

    l1_total, l2_total = l1['total'], l2['total']
    l1_eq,    l2_eq    = l1['equipment'], l2['equipment']
    l1_in,    l1_out   = l1['indoor'],    l1['outdoor']
    l2_in,    l2_out   = l2['indoor'],    l2['outdoor']

    top_fault = fault_list[0][0] if fault_list else ''
    l1_faults, l2_faults = l1['faults'], l2['faults']
    l2_top    = list(l2_faults.keys())[0] if l2_faults else ''

    l1_months = list(l1['monthly'].keys())
    l1_vals   = list(l1['monthly'].values())
    l2_months = list(l2['monthly'].keys())
    l2_vals   = list(l2['monthly'].values())

    monthly_trend = ' → '.join(f"{k} {v}건" for k, v in l1['monthly'].items())

    # 역별 상위
    l1_st = list(l1['stations'].items())[:6]
    l2_st = list(l2['stations'].items())[:5]
    # 1호선/2호선 각각 독립 라벨·값 (역사명이 서로 다름)
    bar_labels  = [s[0] for s in l1_st]   # 1호선 역사명 (x축)
    bar_l1      = [s[1] for s in l1_st]   # 1호선 건수
    bar_l2      = [l2['stations'].get(s[0], 0) for s in l1_st]  # 구버전 호환용(미사용)
    # 2호선 독립 데이터
    bar_labels2 = [s[0] for s in l2_st]   # 2호선 역사명
    bar_l2_vals = [s[1] for s in l2_st]   # 2호선 건수

    def st_table(items, total_n, hdr_col):
        rows = []
        hdr  = (
            f'[{{text:"역사",options:{{bold:true,color:"FFFFFF",fill:{{color:"{hdr_col}"}}}}}}, '
            f'{{text:"건수",options:{{bold:true,color:"FFFFFF",fill:{{color:"{hdr_col}"}}}}}}, '
            f'{{text:"구성비",options:{{bold:true,color:"FFFFFF",fill:{{color:"{hdr_col}"}}}}}}]'
        )
        for name, cnt in items:
            rows.append(f'["{name}","{cnt}","{_pct(cnt,total_n)}%"]')
        rows.append(
            f'[{{text:"합  계",options:{{bold:true}}}}, '
            f'{{text:"{total_n}",options:{{bold:true}}}}, '
            f'{{text:"100%",options:{{bold:true}}}}]'
        )
        return '[' + hdr + ',\n' + ',\n'.join(rows) + ']'

    st1_js  = st_table(l1_st, l1_total, sec)
    st2_js  = st_table(l2_st, l2_total, ter)
    st1_h   = round(0.35 + (len(l1_st)+1)*0.28, 2)
    st2_h   = round(0.35 + (len(l2_st)+1)*0.28, 2)

    # 장치별 원인분석 슬라이드
    fault_slides = _fault_slides(fault_list, slide_bg, pri, is_el)
    conclusion   = _conclusion_cards(fault_list[:5], pri)

    io_l1 = '옥외형이 옥내형보다 높음' if l1_out > l1_in else '옥내형이 옥외형보다 높음'
    io_l2 = '옥외형이 옥내형보다 높음' if l2_out > l2_in else '옥내형이 옥외형보다 높음'

    # output_path: Windows 경로는 역슬래시가 있으므로 JSON으로 안전하게 직렬화
    out_json = json.dumps(output_path)   # → "C:\\Users\\...\\xxx.pptx"

    # ── 목차 동적 페이지 번호 계산
    import math as _math
    _n_fault_slides = max(1, _math.ceil(len(fault_list) / 4))
    _p_fault_ov     = 6
    _p_io           = 6 + _n_fault_slides + 1
    _p_conc         = 6 + _n_fault_slides + 2
    # 신규 슬라이드 페이지 번호 (순서: 옥내외→전년비교→반복→결론)
    _has_yoy        = yoy is not None
    _yoy_data       = (yoy or {}).get(rt.lower(), {})
    _src_data       = stats['el'] if rt == 'EL' else stats['es']
    _repeats        = (_src_data.get('repeats') or [])
    _has_repeats    = len(_repeats) > 0
    # 페이지 번호: 옥내외(_p_io) 다음부터
    _p_yoy     = _p_io + 1 if _has_yoy else None
    _yoy_pages = 2 if _has_yoy else 0                          # 전년비교는 2슬라이드
    _p_repeat  = (_p_io + 1 + _yoy_pages) if _has_repeats else None
    # 결론은 항상 마지막
    _p_conc    = _p_io + _yoy_pages + (1 if _has_repeats else 0) + 1

    # 결론 슬라이드 제목 번호 동적 계산
    _conc_num  = (7 if _has_yoy and _has_repeats
                  else 6 if _has_yoy or _has_repeats
                  else 5)
    conc_title_js = f"'{_conc_num}) 종합 결론 및 향후 관리 방향'"

    # 반복장애 제목 번호
    _repeat_num = 6 if _has_yoy else 5

    lines = [
        "const pptxgen = require('pptxgenjs');",
        "",
        f"const PRI={_q(pri)}, ACC={_q(acc)}, SEC={_q(sec)}, TER={_q(ter)};",
        f"const SIL='D4DCE8', WHT='FFFFFF', GRY='64748B';",
        f"const SBG={_q(slide_bg)}, FTC={_q(foot_c)};",
        "",
        "const ms = () => ({type:'outer',color:'000000',blur:6,offset:3,angle:135,opacity:0.12});",
        "",
        "function tb(sl,title,sub){",
        f"  sl.addShape('rect',{{x:0.4,y:0.22,w:0.06,h:0.55,fill:{{color:ACC}},line:{{color:ACC}}}});",
        "  sl.addText(title,{x:0.55,y:0.18,w:8.5,h:0.45,fontSize:22,bold:true,color:PRI,fontFace:'Arial',margin:0,valign:'middle'});",
        "  if(sub) sl.addText(sub,{x:0.55,y:0.62,w:8.5,h:0.28,fontSize:11,color:GRY,fontFace:'Arial',margin:0});",
        "  sl.addShape('rect',{x:0.4,y:0.9,w:9.2,h:0.02,fill:{color:SIL},line:{color:SIL}});",
        "}",
        "function ft(sl,label){",
        f"  sl.addShape('rect',{{x:0,y:5.35,w:10,h:0.275,fill:{{color:PRI}},line:{{color:PRI}}}});",
        f"  sl.addText('대구교통공사 | 승강설비 장애분석 보고서',{{x:0.3,y:5.36,w:7,h:0.24,fontSize:9,color:FTC,fontFace:'Arial',margin:0}});",
        "  sl.addText(label,{x:7.3,y:5.36,w:2.4,h:0.24,fontSize:9,color:FTC,fontFace:'Arial',align:'right',margin:0});",
        "}",
        "function kpi(sl,x,y,w,h,t,v,s,c){",
        "  sl.addShape('rect',{x,y,w,h,fill:{color:WHT},line:{color:SIL,pt:1},shadow:ms()});",
        "  sl.addShape('rect',{x,y,w,h:0.07,fill:{color:c},line:{color:c}});",
        "  sl.addText(t,{x:x+0.12,y:y+0.12,w:w-0.24,h:0.28,fontSize:10,color:GRY,fontFace:'Arial',margin:0});",
        "  sl.addText(v,{x:x+0.08,y:y+0.38,w:w-0.16,h:0.65,fontSize:27,bold:true,color:PRI,fontFace:'Arial Black',align:'center',margin:0});",
        "  sl.addText(s,{x:x+0.08,y:y+0.98,w:w-0.16,h:0.22,fontSize:9,color:GRY,fontFace:'Arial',align:'center',margin:0});",
        "}",
        "",
        "const pres = new pptxgen();",
        "pres.layout = 'LAYOUT_16x9';",
        f"pres.title = '{title_kr} 장애분석 보고서';",
        "",
        "// ── 표지",
        "{ const sl = pres.addSlide();",
        f"  sl.background = {{color:'{cover_bg}'}};",
        f"  sl.addShape('rect',{{x:0,y:0,w:10,h:1.2,fill:{{color:'{top_bar}'}},line:{{color:'{top_bar}'}}}});",
        "  sl.addShape('rect',{x:8.4,y:1.4,w:0.04,h:2.8,fill:{color:ACC},line:{color:ACC}});",
        f"  sl.addShape('ellipse',{{x:0.6,y:1.8,w:1.2,h:1.2,fill:{{color:ACC,transparency:85}},line:{{color:ACC,transparency:50}}}});",
        f"  sl.addText('{label_cd}',{{x:0.6,y:1.9,w:1.2,h:1.0,fontSize:22,bold:true,color:ACC,align:'center',fontFace:'Arial Black',margin:0}});",
        f"  sl.addText('{title_kr} 장애분석 보고서',{{x:0.5,y:1.5,w:7.5,h:0.7,fontSize:30,bold:true,color:'FFFFFF',fontFace:'Arial',margin:0}});",
        f"  sl.addText('{title_en} Fault Analysis Report',{{x:0.5,y:2.15,w:7.5,h:0.4,fontSize:14,color:'88AACC',fontFace:'Arial',margin:0}});",
        "  sl.addShape('rect',{x:0.5,y:2.6,w:3.5,h:0.04,fill:{color:ACC},line:{color:ACC}});",
        f"  sl.addText('{year}년 {quarter} ({period})',{{x:0.5,y:2.75,w:6,h:0.4,fontSize:13,color:'AABBD0',fontFace:'Arial',margin:0}});",
        "  sl.addText('대구교통공사 | 기계팀',{x:0.5,y:5.05,w:7,h:0.3,fontSize:11,color:'AABBD0',fontFace:'Arial',margin:0});",
        f"  [{_j(['총 장애건수', f'{total}건'])},{_j(['1호선', f'{l1_total}건'])},{_j(['2호선', f'{l2_total}건'])},{_j(['평균복구', f'{avg_rec}분'])}].forEach(([t,v],i)=>{{",
        "    const sx=0.5+i*1.85;",
        f"    sl.addShape('rect',{{x:sx,y:3.35,w:1.65,h:1.15,fill:{{color:'1E3A5F'}},line:{{color:TER,pt:1}},shadow:ms()}});",
        "    sl.addText(t,{x:sx,y:3.42,w:1.65,h:0.28,fontSize:9,color:'88AACC',align:'center',fontFace:'Arial',margin:0});",
        "    sl.addText(v,{x:sx,y:3.68,w:1.65,h:0.52,fontSize:22,bold:true,color:'FFFFFF',align:'center',fontFace:'Arial Black',margin:0});",
        "  });",
        "}",
        "",
        "// ── 목차",
        "{ const sl = pres.addSlide();",
        f"  sl.background = {{color:SBG}};",
        f"  sl.addShape('rect',{{x:0,y:0,w:10,h:5.625,fill:{{color:SBG}},line:{{color:SBG}}}});",
        f"  sl.addShape('rect',{{x:0,y:0,w:0.08,h:5.625,fill:{{color:ACC}},line:{{color:ACC}}}});",
        f"  sl.addShape('rect',{{x:0,y:0,w:10,h:1.1,fill:{{color:PRI}},line:{{color:PRI}}}});",
        f"  sl.addText('목  차',{{x:0.3,y:0.15,w:5,h:0.5,fontSize:28,bold:true,color:'FFFFFF',fontFace:'Arial',margin:0}});",
        f"  sl.addText('Contents',{{x:0.3,y:0.65,w:5,h:0.35,fontSize:14,color:'88AACC',fontFace:'Arial',margin:0}});",
        f"  sl.addShape('rect',{{x:0,y:5.35,w:10,h:0.275,fill:{{color:PRI}},line:{{color:PRI}}}});",
        f"  sl.addText('대구교통공사 | 승강설비 장애분석 보고서',{{x:0.3,y:5.36,w:7,h:0.24,fontSize:9,color:FTC,fontFace:'Arial',margin:0}});",
        f"  sl.addText('{title_kr} 장애분석 보고서',{{x:7.3,y:5.36,w:2.4,h:0.24,fontSize:9,color:FTC,fontFace:'Arial',align:'right',margin:0}});",
        f"  const tocItems = [",
        "    ['01', '월별 장애현황 및 감소대책', '1호선·2호선 월별 장애 추이 분석', 'P.3'],",
        "    ['02', '역별 장애현황 및 감소대책', '호선별 역사별 장애 집중도 분석', 'P.4'],",
        f"    ['03', '장치별 장애현황 및 감소대책', '장애유형별 원인분석 및 예방 대책', 'P.{_p_fault_ov}'],",
        f"    ['04', '옥내/옥외형 장애현황', '설치환경에 따른 장애율 비교 분석', 'P.{_p_io}'],",
        *([ f"    ['05', '전년 동기 비교 분석', '{_yoy_data.get('year_last','전년')}년 {quarter} 대비 변화', 'P.{_p_yoy}']," ] if _has_yoy else []),
        *([ f"    ['{'06' if _has_yoy else '05'}', '반복장애 예방점검 권고', '분기 내 반복 장애 장치 집중 관리', 'P.{_p_repeat}']," ] if _has_repeats else []),
        f"    ['{('07' if _has_yoy and _has_repeats else '06' if _has_yoy or _has_repeats else '05')}', '종합 결론 및 향후 대책', '핵심 개선방향 및 예방관리 방침', 'P.{_p_conc}'],",
        "  ];",
        "  // 항목 수에 따라 높이·간격 자동 조정 (최대 7개까지 슬라이드 안에 수용)",
        "  const nItems = tocItems.length;",
        "  const itemH  = nItems <= 5 ? 0.68 : nItems === 6 ? 0.60 : 0.52;",
        "  const gap    = nItems <= 5 ? 0.77 : nItems === 6 ? 0.67 : 0.585;",
        "  const yStart = nItems <= 5 ? 1.15 : 1.06;",
        "  const fT     = nItems <= 5 ? 12   : 11;",
        "  const fD     = nItems <= 5 ? 9.5  : 8.5;",
        "  tocItems.forEach(([num, title, desc, page], i) => {",
        "    const y = yStart + i * gap;",
        "    sl.addShape('rect',{x:0.2,y:y,w:9.6,h:itemH,fill:{color:'FFFFFF'},line:{color:SIL,pt:1}});",
        "    sl.addShape('rect',{x:0.2,y:y,w:0.7,h:itemH,fill:{color:ACC},line:{color:ACC}});",
        "    sl.addText(num,{x:0.2,y:y+itemH*0.22,w:0.7,h:itemH*0.56,fontSize:12,bold:true,color:'FFFFFF',align:'center',fontFace:'Arial Black',margin:0});",
        "    sl.addText(title,{x:1.05,y:y+0.05,w:7.4,h:itemH*0.45,fontSize:fT,bold:true,color:PRI,fontFace:'Arial',margin:0});",
        "    sl.addText(desc,{x:1.05,y:y+itemH*0.52,w:7.4,h:itemH*0.42,fontSize:fD,color:GRY,fontFace:'Arial',margin:0});",
        "    sl.addText(page,{x:8.8,y:y+itemH*0.18,w:0.9,h:itemH*0.5,fontSize:11,bold:true,color:ACC,fontFace:'Arial',align:'right',margin:0});",
        "  });",
        "}",
        "",
        "// ── 월별 장애현황",
        "{ const sl = pres.addSlide();",
        f"  sl.background = {{color:SBG}};",
        f"  tb(sl,'1) 월별 장애현황 및 감소대책','{quarter} 월별 장애건수 추이 분석');",
        f"  ft(sl,'월별 현황');",
        f"  kpi(sl,0.4,1.05,1.9,1.35,'{quarter} 총 장애','{total}건','1·2호선 합계',PRI);",
        f"  kpi(sl,2.4,1.05,1.9,1.35,'1호선','{l1_total}건','장비 {l1_eq}대 운용',SEC);",
        f"  kpi(sl,4.4,1.05,1.9,1.35,'2호선','{l2_total}건','장비 {l2_eq}대 운용',TER);",
        f"  kpi(sl,6.4,1.05,1.9,1.35,'평균 복구','{avg_rec}분','전호선 평균','E8843A');",
        # 최다 장애 KPI 카드 - '1호선 1위' 라벨 삭제, 글자크기 줄여 한 줄 표시
        f"  sl.addShape('rect',{{x:8.3,y:1.05,w:1.3,h:1.35,fill:{{color:'FFFFFF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});",
        f"  sl.addShape('rect',{{x:8.3,y:1.05,w:1.3,h:0.07,fill:{{color:'C0392B'}},line:{{color:'C0392B'}}}});",
        f"  sl.addText('최다 장애',{{x:8.42,y:1.17,w:1.06,h:0.28,fontSize:10,color:GRY,fontFace:'Arial',margin:0}});",
        f"  sl.addText('{top_fault}',{{x:8.3,y:1.38,w:1.3,h:0.58,fontSize:16,bold:true,color:'C0392B',fontFace:'Arial Black',align:'center',margin:0,valign:'middle'}});",
        f"  sl.addChart(pres.charts.BAR,[{{name:'1호선',labels:{_j(l1_months)},values:{_j(l1_vals)}}},{{name:'2호선',labels:{_j(l2_months)},values:{_j(l2_vals)}}}],",
        "    {x:0.4,y:2.55,w:5.5,h:2.6,barDir:'col',barGrouping:'clustered',",
        "     chartColors:[SEC,TER],chartArea:{fill:{color:'FFFFFF'},roundedCorners:true},",
        "     catAxisLabelColor:GRY,valAxisLabelColor:GRY,",
        "     valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},",
        "     showValue:true,dataLabelColor:PRI,showLegend:true,legendPos:'b',legendFontSize:10,",
        f"     showTitle:true,title:'호선별 월별 장애건수 (건)',titleFontSize:11,titleColor:PRI}});",
        f"  sl.addShape('rect',{{x:6.1,y:2.55,w:3.5,h:2.6,fill:{{color:'FFFFFF'}},line:{{color:SIL,pt:1}},shadow:ms()}});",
        f"  sl.addShape('rect',{{x:6.1,y:2.55,w:3.5,h:0.38,fill:{{color:PRI}},line:{{color:PRI}}}});",
        "  sl.addText('분석 의견',{x:6.2,y:2.56,w:3.3,h:0.36,fontSize:11,bold:true,color:'FFFFFF',fontFace:'Arial',margin:0});",
        "  sl.addText(" + _build_analysis_opinion(quarter, l1_total, l2_total, monthly_trend, top_fault, l2_top, l1_faults, l2_faults, l1_months, l1_vals) + ",",
        "    {x:6.15,y:2.98,w:3.4,h:2.1,fontSize:9,color:PRI,fontFace:'Arial',valign:'top',margin:5});",
        "}",
        "",
        "// ── 역별 장애현황 (테이블)",
        "{ const sl = pres.addSlide();",
        f"  sl.background = {{color:SBG}};",
        f"  tb(sl,'2) 역별 장애현황 및 감소대책','장애 다발 역사 파악 및 집중 관리 방향 수립');",
        f"  ft(sl,'역별 현황');",
        f"  sl.addShape('rect',{{x:0.4,y:1.0,w:4.4,h:0.35,fill:{{color:SEC}},line:{{color:SEC}}}});",
        f"  sl.addText('1호선 역별 장애현황 (상위 {len(l1_st)}개역)',{{x:0.4,y:1.0,w:4.4,h:0.35,fontSize:11,bold:true,color:'FFFFFF',align:'center',margin:0}});",
        f"  sl.addTable({st1_js},{{x:0.4,y:1.35,w:4.4,h:{st1_h},border:{{pt:1,color:SIL}},fill:{{color:'FFFFFF'}},colW:[1.8,1.3,1.3],fontFace:'Arial',fontSize:11,align:'center',rowH:0.28}});",
        f"  sl.addShape('rect',{{x:5.2,y:1.0,w:4.4,h:0.35,fill:{{color:TER}},line:{{color:TER}}}});",
        f"  sl.addText('2호선 역별 장애현황 (상위 {len(l2_st)}개역)',{{x:5.2,y:1.0,w:4.4,h:0.35,fontSize:11,bold:true,color:'FFFFFF',align:'center',margin:0}});",
        f"  sl.addTable({st2_js},{{x:5.2,y:1.35,w:4.4,h:{st2_h},border:{{pt:1,color:SIL}},fill:{{color:'FFFFFF'}},colW:[1.8,1.3,1.3],fontFace:'Arial',fontSize:11,align:'center',rowH:0.28}});",
        "}",
        "",
        "// ── 역별 장애현황 (막대그래프 - 1호선/2호선 각각 표시)",
        "{ const sl = pres.addSlide();",
        f"  sl.background = {{color:SBG}};",
        f"  tb(sl,'2) 역별 장애현황 및 감소대책 (그래프)','호선별 역사별 장애건수 비교');",
        f"  ft(sl,'역별 현황');",
        f"  sl.addShape('rect',{{x:0.3,y:0.95,w:4.7,h:0.3,fill:{{color:SEC}},line:{{color:SEC}}}});",
        f"  sl.addText('1호선 역별 장애건수',{{x:0.3,y:0.95,w:4.7,h:0.3,fontSize:11,bold:true,color:'FFFFFF',align:'center',margin:0}});",
        f"  sl.addChart(pres.charts.BAR,[{{name:'1호선',labels:{_j(bar_labels)},values:{_j(bar_l1)}}}],",
        "    {x:0.3,y:1.25,w:4.7,h:3.85,barDir:'bar',barGrouping:'clustered',",
        "     chartColors:[SEC],chartArea:{fill:{color:'FFFFFF'},roundedCorners:true},",
        "     catAxisLabelColor:GRY,valAxisLabelColor:GRY,catAxisFontSize:10,",
        "     valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},",
        "     showValue:true,dataLabelFontSize:10,dataLabelColor:PRI,",
        f"     showLegend:false}});",
        f"  sl.addShape('rect',{{x:5.2,y:0.95,w:4.7,h:0.3,fill:{{color:TER}},line:{{color:TER}}}});",
        f"  sl.addText('2호선 역별 장애건수',{{x:5.2,y:0.95,w:4.7,h:0.3,fontSize:11,bold:true,color:'FFFFFF',align:'center',margin:0}});",
        f"  sl.addChart(pres.charts.BAR,[{{name:'2호선',labels:{_j(bar_labels2)},values:{_j(bar_l2_vals)}}}],",
        "    {x:5.2,y:1.25,w:4.7,h:3.85,barDir:'bar',barGrouping:'clustered',",
        "     chartColors:[TER],chartArea:{fill:{color:'FFFFFF'},roundedCorners:true},",
        "     catAxisLabelColor:GRY,valAxisLabelColor:GRY,catAxisFontSize:10,",
        "     valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},",
        "     showValue:true,dataLabelFontSize:10,dataLabelColor:PRI,",
        f"     showLegend:false}});",
        "}",
        "",
        "// ── 장치별 개요",
        "{ const sl = pres.addSlide();",
        f"  sl.background = {{color:SBG}};",
        f"  tb(sl,'3) 장치별 장애현황 및 감소대책 (개요)','호선별 장치유형별 장애분포 분석');",
        f"  ft(sl,'장치별 현황');",
        f"  sl.addChart(pres.charts.DOUGHNUT,[{{name:'1호선 장치별',labels:{_j(list(l1_faults.keys()))},values:{_j(list(l1_faults.values()))}}}],",
        "    {x:0.3,y:1.0,w:4.8,h:3.5,chartColors:['C0392B','E8843A','2C5F8A','0D7A8A','F5A623','27AE60','64748B','8E44AD','16A085','2AAF7E'],",
        f"     showPercent:true,showLegend:true,legendPos:'b',legendFontSize:8,showTitle:true,title:'1호선 장치별 장애 구성비',titleFontSize:11,titleColor:PRI,",
        "     chartArea:{fill:{color:'FFFFFF'},roundedCorners:true},dataLabelColor:'FFFFFF'});",
        f"  sl.addChart(pres.charts.DOUGHNUT,[{{name:'2호선 장치별',labels:{_j(list(l2_faults.keys()))},values:{_j(list(l2_faults.values()))}}}],",
        "    {x:5.1,y:1.0,w:4.8,h:3.5,chartColors:['C0392B','E8843A','2C5F8A','0D7A8A','F5A623','27AE60','64748B','8E44AD','16A085','2AAF7E'],",
        f"     showPercent:true,showLegend:true,legendPos:'b',legendFontSize:8,showTitle:true,title:'2호선 장치별 장애 구성비',titleFontSize:11,titleColor:PRI,",
        "     chartArea:{fill:{color:'FFFFFF'},roundedCorners:true},dataLabelColor:'FFFFFF'});",
        f"  sl.addShape('rect',{{x:0.4,y:4.65,w:9.2,h:0.55,fill:{{color:'EEF2F7'}},line:{{color:SIL,pt:1}}}});",
        "  sl.addText(" + _q('⇒ 1·2호선 공통 최다 장애 : ' + top_fault + '  /  장치별 원인분석 및 감소대책 후속 슬라이드 참고') + ",",
        "    {x:0.55,y:4.68,w:9.0,h:0.48,fontSize:10,bold:true,color:PRI,fontFace:'Arial',margin:0,valign:'middle'});",
        "}",
        "",
        fault_slides,
        "",
        "// ── 옥내/옥외 현황",
        "{ const sl = pres.addSlide();",
        f"  sl.background = {{color:SBG}};",
        f"  tb(sl,'4) 옥내/옥외형 장애현황 및 감소대책','설치환경에 따른 장애율 비교 분석');",
        f"  ft(sl,'옥내/옥외 현황');",
        f"  sl.addChart(pres.charts.BAR,[{{name:'옥내형',labels:['1호선','2호선'],values:[{l1_in},{l2_in}]}},{{name:'옥외형',labels:['1호선','2호선'],values:[{l1_out},{l2_out}]}}],",
        "    {x:0.4,y:1.0,w:5.2,h:2.8,barDir:'col',barGrouping:'clustered',",
        "     chartColors:[TER,'E8843A'],chartArea:{fill:{color:'FFFFFF'},roundedCorners:true},",
        "     catAxisLabelColor:GRY,valAxisLabelColor:GRY,",
        "     valGridLine:{color:'E2E8F0',size:0.5},catGridLine:{style:'none'},",
        "     showValue:true,dataLabelColor:PRI,showLegend:true,legendPos:'b',legendFontSize:10,",
        f"     showTitle:true,title:'호선별 옥내/옥외형 장애건수',titleFontSize:11,titleColor:PRI}});",
        "  sl.addTable([",
        f"    [{{text:'구  분',options:{{bold:true,color:'FFFFFF',fill:{{color:PRI}}}}}},{{text:'옥내형',options:{{bold:true,color:'FFFFFF',fill:{{color:PRI}}}}}},{{text:'옥외형',options:{{bold:true,color:\'FFFFFF\',fill:{{color:PRI}}}}}},{{text:'합계',options:{{bold:true,color:'FFFFFF',fill:{{color:PRI}}}}}}],",
        f"    ['1호선','{l1_in}건 ({_pct(l1_in,l1_total)}%)','{l1_out}건 ({_pct(l1_out,l1_total)}%)','{l1_total}건'],",
        f"    ['2호선','{l2_in}건 ({_pct(l2_in,l2_total)}%)','{l2_out}건 ({_pct(l2_out,l2_total)}%)','{l2_total}건'],",
        f"    [{{text:'합  계',options:{{bold:true}}}},{{text:'{l1_in+l2_in}건',options:{{bold:true}}}},{{text:'{l1_out+l2_out}건',options:{{bold:true}}}},{{text:'{total}건',options:{{bold:true}}}}],",
        "  ],{x:5.7,y:1.0,w:3.6,h:1.4,border:{pt:1,color:SIL},fill:{color:'FFFFFF'},colW:[0.8,1.1,1.1,0.6],fontFace:'Arial',fontSize:10,align:'center',rowH:0.32});",
        f"  sl.addShape('rect',{{x:5.7,y:2.6,w:3.6,h:1.2,fill:{{color:'FFFFFF'}},line:{{color:SIL,pt:1}},shadow:ms()}});",
        f"  sl.addShape('rect',{{x:5.7,y:2.6,w:3.6,h:0.32,fill:{{color:TER}},line:{{color:TER}}}});",
        "  sl.addText('분석 의견',{x:5.8,y:2.61,w:3.5,h:0.3,fontSize:10,bold:true,color:'FFFFFF',margin:0});",
        "  sl.addText(" + _q("▶ 1호선 : " + io_l1 + "\n\n▶ 2호선 : " + io_l2 + "\n\n⇒ 설치 환경별 집중 점검 체계 강화 필요") + ",",
        "    {x:5.8,y:2.95,w:3.5,h:0.82,fontSize:9.5,color:PRI,fontFace:'Arial',margin:3,valign:'top'});",
        f"  sl.addShape('rect',{{x:0.4,y:4.0,w:9.2,h:1.12,fill:{{color:'FFFFFF'}},line:{{color:SIL,pt:1}},shadow:ms()}});",
        f"  sl.addShape('rect',{{x:0.4,y:4.0,w:0.06,h:1.12,fill:{{color:ACC}},line:{{color:ACC}}}});",
        "  sl.addText('옥내·옥외형 장애 감소대책',{x:0.6,y:4.06,w:5,h:0.28,fontSize:11,bold:true,color:PRI,margin:0});",
        "  sl.addText('① 옥외형 : 방수·방진 커버 상태 점검, 우천·강풍 후 즉시 스위치류 점검 실시\\n② 옥내형 : 승강장 내 분진·습기 차단 관리, 스텝-스커트 간격 정기 확인\\n③ 공통 : 계절별 환경 변화에 맞는 점검 항목 세분화 및 이상 징후 조기 감지',",
        "    {x:0.6,y:4.33,w:9.0,h:0.76,fontSize:10,color:PRI,fontFace:'Arial',margin:0,valign:'top'});",
        "}",
        "",
        "// ── 종합 결론",
        "{ const sl = pres.addSlide();",
        f"  sl.background = {{color:SBG}};",
        f"  tb(sl,{conc_title_js},'핵심 개선과제 및 예방관리 방침');",
        f"  ft(sl,'종합 결론');",
        conclusion,
        f"  sl.addShape('rect',{{x:5.25,y:{1.05+2*1.45:.2f},w:4.65,h:1.3,fill:{{color:PRI}},line:{{color:PRI}},shadow:ms()}});",
        f"  sl.addText('\"지속적인 예방 관리를 통해\\n승객 안전과 서비스 품질을 높이겠습니다.\"',",
        f"    {{x:5.35,y:{1.05+2*1.45+0.25:.2f},w:4.45,h:0.8,fontSize:12,bold:true,color:'FFFFFF',align:'center',fontFace:'Arial',margin:0,italic:true}});",
        "}",
        "",
        f"pres.writeFile({{fileName:{out_json}}}).then(()=>console.log('saved'));",
    ]

    # ── 순서: 전년비교 → 반복장애 → 종합결론(마지막) → writeFile
    # "// ── 종합 결론" 항목 바로 앞 인덱스 찾아서 삽입
    conc_idx = next(
        (i for i, item in enumerate(lines) if '// ── 종합 결론' in str(item)),
        len(lines) - 2
    )

    if _has_yoy and _yoy_data:
        yoy_slide = _yoy_slide(_yoy_data, slide_bg, pri, sec, ter, acc, quarter)
        lines.insert(conc_idx, yoy_slide)
        conc_idx += 1   # 삽입 후 결론 인덱스 밀림

    if _has_repeats:
        repeat_slide = _repeat_slide(_repeats, slide_bg, pri, acc, title_kr, _repeat_num)
        lines.insert(conc_idx, repeat_slide)
        # 최종 순서: ...옥내외 → 전년비교 → 반복장애 → 종합결론 → writeFile

    return '\n'.join(lines)


def _esc(s: str) -> str:
    return str(s).replace('\\', '\\\\').replace("'", "\\'").replace('\n', '\\n')


def _fault_slides(fault_list, slide_bg, pri, is_el) -> str:
    items  = list(fault_list)
    if not items:
        return ''

    groups = []
    i = 0
    while i < len(items):
        groups.append(items[i:i+4])
        i += 4

    blocks = []
    for gi, group in enumerate(groups):
        part  = f"{gi+1}/{len(groups)}"
        names = ' · '.join(k for k, _ in group)
        cards = []

        if len(group) <= 3:
            for ci, (key, info) in enumerate(group):
                cause, action, measure, color = info
                x = 0.4 + ci * 3.12
                cards.append(
                    f"  sl.addShape('rect',{{x:{x},y:1.0,w:3.0,h:3.7,fill:{{color:'FFFFFF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});\n"
                    f"  sl.addShape('rect',{{x:{x},y:1.0,w:3.0,h:0.42,fill:{{color:'{color}'}},line:{{color:'{color}'}}}});\n"
                    f"  sl.addText({_q(key+' 관련 장애')},{{x:{x+0.1},y:1.06,w:2.8,h:0.3,fontSize:11,bold:true,color:'FFFFFF',fontFace:'Arial',margin:0}});\n"
                    f"  sl.addText('【원인】',{{x:{x+0.12},y:1.52,w:2.8,h:0.25,fontSize:10,bold:true,color:'{pri}',margin:0}});\n"
                    f"  sl.addText({_q(cause)},{{x:{x+0.12},y:1.74,w:2.76,h:0.72,fontSize:9.5,color:'{pri}',fontFace:'Arial',margin:0,valign:'top'}});\n"
                    f"  sl.addText('【조치내역】',{{x:{x+0.12},y:2.46,w:2.8,h:0.25,fontSize:10,bold:true,color:'64748B',margin:0}});\n"
                    f"  sl.addText({_q(action)},{{x:{x+0.12},y:2.68,w:2.76,h:0.52,fontSize:9.5,color:'64748B',fontFace:'Arial',margin:0,valign:'top'}});\n"
                    f"  sl.addText('【감소대책】',{{x:{x+0.12},y:3.22,w:2.8,h:0.25,fontSize:10,bold:true,color:'0D7A8A',margin:0}});\n"
                    f"  sl.addText({_q(measure)},{{x:{x+0.12},y:3.44,w:2.76,h:0.62,fontSize:9.5,color:'0D7A8A',fontFace:'Arial',margin:0,valign:'top'}});"
                )
        else:
            for ci, (key, info) in enumerate(group):
                cause, action, measure, color = info
                col = ci // 2
                row = ci % 2
                x   = 0.4 + col * 4.8
                y   = 1.0 + row * 2.18
                cards.append(
                    f"  sl.addShape('rect',{{x:{x},y:{y},w:4.6,h:2.04,fill:{{color:'FFFFFF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});\n"
                    f"  sl.addShape('rect',{{x:{x},y:{y},w:4.6,h:0.42,fill:{{color:'{color}'}},line:{{color:'{color}'}}}});\n"
                    f"  sl.addText({_q(key+' 관련 장애')},{{x:{x+0.12},y:{y+0.06},w:4.36,h:0.3,fontSize:12,bold:true,color:'FFFFFF',fontFace:'Arial',margin:0}});\n"
                    f"  sl.addText({_q('【원인】 '+cause)},{{x:{x+0.12},y:{y+0.5},w:4.36,h:0.5,fontSize:9.5,color:'{pri}',fontFace:'Arial',margin:0,valign:'top'}});\n"
                    f"  sl.addText({_q('【조치】 '+action)},{{x:{x+0.12},y:{y+0.98},w:4.36,h:0.36,fontSize:9.5,color:'64748B',fontFace:'Arial',margin:0,valign:'top'}});\n"
                    f"  sl.addText({_q('【대책】 '+measure)},{{x:{x+0.12},y:{y+1.32},w:4.36,h:0.58,fontSize:9.5,color:'0D7A8A',fontFace:'Arial',margin:0,valign:'top'}});"
                )

        blocks.append(
            f"// ── 장치별 원인분석 {part}\n"
            f"{{ const sl = pres.addSlide();\n"
            f"  sl.background = {{color:'{slide_bg}'}};\n"
            f"  tb(sl,'3) 장치별 장애원인 분석 및 대책 ({part})',{_q(names)});\n"
            f"  ft(sl,'원인분석 {gi+1}');\n"
            + '\n'.join(cards) +
            "\n}"
        )

    return '\n\n'.join(blocks)


def _conclusion_cards(top5, pri) -> str:
    lines = []
    for i, (key, info) in enumerate(top5):
        _, _, measure, color = info
        col = 0 if i < 3 else 1
        row = i if i < 3 else i - 3
        x   = 0.4 + col * 4.85
        y   = 1.05 + row * 1.45
        num = f"{i+1:02d}"
        body = _esc(measure[:80])
        lines.append(
            f"  sl.addShape('rect',{{x:{x},y:{y},w:4.65,h:1.3,fill:{{color:'FFFFFF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});\n"
            f"  sl.addShape('rect',{{x:{x},y:{y},w:0.55,h:1.3,fill:{{color:'{color}'}},line:{{color:'{color}'}}}});\n"
            f"  sl.addText('{num}',{{x:{x},y:{y+0.43},w:0.55,h:0.44,fontSize:15,bold:true,color:'FFFFFF',align:'center',fontFace:'Arial Black',margin:0}});\n"
            f"  sl.addText({_q(key+' 집중 관리')},{{x:{x+0.65},y:{y+0.12},w:3.9,h:0.3,fontSize:11.5,bold:true,color:'{pri}',fontFace:'Arial',margin:0}});\n"
            f"  sl.addText('{body}',{{x:{x+0.65},y:{y+0.45},w:3.9,h:0.75,fontSize:9.5,color:'64748B',fontFace:'Arial',margin:0,valign:'top'}});"
        )
    return '\n'.join(lines)


# ── 전년 동기 비교 슬라이드 ────────────────────────────────────────────────────


# ── 전년 동기 비교 슬라이드 (2장) ────────────────────────────────────────────

def _yoy_slide(yoy_data: dict, slide_bg, pri, sec, ter, acc, quarter) -> str:
    """
    전년 동기 비교 분석 슬라이드 2장:
      슬라이드 A: KPI 요약 + 월별 꺾은선 차트
      슬라이드 B: 장치별 비교 차트 + 상세 분석의견
    """
    year_this = yoy_data.get('year_this', 2026)
    year_last = yoy_data.get('year_last', 2025)
    tt  = yoy_data.get('total_this', 0)
    tl  = yoy_data.get('total_last', 0)
    td  = yoy_data.get('total_diff', 0)
    tp  = yoy_data.get('total_pct', 0.0)
    l1  = yoy_data.get('line1', {})
    l2  = yoy_data.get('line2', {})

    # 증감 표시
    def arrow(d): return '▲' if d > 0 else ('▼' if d < 0 else '―')
    def acolor(d): return 'C0392B' if d > 0 else ('27AE60' if d < 0 else '64748B')

    td_arrow  = arrow(td);              td_color  = acolor(td)
    l1d = l1.get('diff', 0);            l1d_arrow = arrow(l1d); l1d_color = acolor(l1d)
    l2d = l2.get('diff', 0);            l2d_arrow = arrow(l2d); l2d_color = acolor(l2d)
    tp_sign   = f"{tp:+.1f}%"

    # 월별
    monthly  = yoy_data.get('monthly', {})
    m_labels = list(monthly.keys())
    m_this   = [v.get('this', 0) for v in monthly.values()]
    m_last   = [v.get('last', 0) for v in monthly.values()]

    # 장치별 상위 5개
    faults    = yoy_data.get('faults', {})
    top_f     = list(faults.items())[:5]
    f_labels  = [k for k, _ in top_f]
    f_this    = [v.get('this', 0) for _, v in top_f]
    f_last    = [v.get('last', 0) for _, v in top_f]

    # 주의 역사 (증가 상위 3)
    stations  = yoy_data.get('stations', {})
    warn_st   = sorted(
        [(s, v) for s, v in stations.items() if v.get('diff', 0) > 0],
        key=lambda x: -x[1].get('diff', 0)
    )[:3]

    # 전체 추세 판정
    trend     = '감소' if td < 0 else ('증가' if td > 0 else '보합')
    trend_clr = '27AE60' if td < 0 else ('C0392B' if td > 0 else '64748B')

    # 장치별 증가/감소 분석
    f_increased = [(k, v) for k, v in faults.items() if v.get('diff', 0) > 0][:3]
    f_decreased = [(k, v) for k, v in faults.items() if v.get('diff', 0) < 0][:3]
    inc_str = ', '.join(f"{k}({v['diff']:+d}건)" for k, v in f_increased) if f_increased else '없음'
    dec_str = ', '.join(f"{k}({v['diff']:+d}건)" for k, v in f_decreased) if f_decreased else '없음'

    warn_str = ', '.join(f"{s}({v['diff']:+d}건)" for s, v in warn_st) if warn_st else '없음'

    # 월별 추이 문자열
    m_trend_parts = [f"{m} {v['this']}건(전년 {v['last']}건)" for m, v in list(monthly.items())]
    m_trend_str   = ' → '.join(m_trend_parts)

    # ──────────────────────────────────────────────────────────────────────────
    # 공통 헬퍼 함수 (JS 인라인)
    header_js = f"""
  sl.addShape('rect',{{x:0.4,y:0.22,w:0.06,h:0.55,fill:{{color:'{acc}'}},line:{{color:'{acc}'}}}});
  sl.addText('5) 전년 동기 비교 분석',{{x:0.55,y:0.18,w:8.5,h:0.45,fontSize:22,bold:true,color:'{pri}',fontFace:'Arial',margin:0,valign:'middle'}});
  sl.addShape('rect',{{x:0.4,y:0.9,w:9.2,h:0.02,fill:{{color:'D4DCE8'}},line:{{color:'D4DCE8'}}}});"""

    # ── 슬라이드 A: KPI + 월별 차트 ──────────────────────────────────────────
    slide_a = f"""
// ── 전년 동기 비교 (슬라이드 1/2): KPI 요약 + 월별 추이
{{ const sl = pres.addSlide();
  sl.background = {{color:'{slide_bg}'}};
  {header_js}
  sl.addText('{year_last}년 {quarter} 대비 {year_this}년 {quarter} 종합 비교',{{x:0.55,y:0.62,w:8.5,h:0.28,fontSize:11,color:'64748B',fontFace:'Arial',margin:0}});
  ft(sl,'전년 비교 1/2');

  // KPI 카드 3개
  const kpis=[
    ['{year_this}년 합계','{tt}건','전년 {tl}건','{td_arrow}{abs(td)}건','{tp_sign}','{td_color}'],
    ['1호선','{l1.get("this",0)}건','전년 {l1.get("last",0)}건','{l1d_arrow}{abs(l1d)}건','{l1.get("pct",0.0):+.1f}%','{l1d_color}'],
    ['2호선','{l2.get("this",0)}건','전년 {l2.get("last",0)}건','{l2d_arrow}{abs(l2d)}건','{l2.get("pct",0.0):+.1f}%','{l2d_color}'],
  ];
  kpis.forEach(([t,v,comp,chg,pct,chgC],i)=>{{
    const x=0.4+i*3.1;
    sl.addShape('rect',{{x,y:1.08,w:2.9,h:1.35,fill:{{color:'FFFFFF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});
    sl.addShape('rect',{{x,y:1.08,w:2.9,h:0.07,fill:{{color:'{pri}'}},line:{{color:'{pri}'}}}});
    sl.addText(t,{{x:x+0.12,y:1.17,w:2.66,h:0.24,fontSize:11,color:'64748B',fontFace:'Arial',margin:0}});
    sl.addText(v,{{x:x+0.08,y:1.4,w:2.74,h:0.5,fontSize:28,bold:true,color:'{pri}',fontFace:'Arial Black',align:'center',margin:0}});
    sl.addText(comp,{{x:x+0.12,y:1.9,w:1.2,h:0.22,fontSize:9.5,color:'888888',fontFace:'Arial',margin:0}});
    sl.addText(chg+' ('+pct+')',{{x:x+1.1,y:1.9,w:1.7,h:0.22,fontSize:10,bold:true,color:chgC,fontFace:'Arial',margin:0}});
  }});

  // 전체 추세 배너
  sl.addShape('rect',{{x:0.4,y:2.6,w:9.2,h:0.38,fill:{{color:'{trend_clr}'}},line:{{color:'{trend_clr}'}}}});
  sl.addShape('rect',{{x:0.4,y:2.6,w:0.06,h:0.38,fill:{{color:'FFFFFF'}},line:{{color:'FFFFFF'}}}});
  sl.addText('전체 추세 : 전년 동기 대비 {trend}  ({td:+d}건, {tp_sign})',{{x:0.6,y:2.63,w:8.8,h:0.32,fontSize:12,bold:true,color:'FFFFFF',fontFace:'Arial',margin:0,valign:'middle'}});

  // 월별 꺾은선 차트 (전체 너비 활용)
  sl.addChart(pres.charts.LINE,
    [{{name:'{year_this}년',labels:{_j(m_labels)},values:{_j(m_this)}}},
     {{name:'{year_last}년',labels:{_j(m_labels)},values:{_j(m_last)}}}],
    {{x:0.4,y:3.12,w:5.8,h:2.18,
      chartColors:['{sec}','BBCCDD'],
      lineSize:2.5,
      chartArea:{{fill:{{color:'FFFFFF'}},roundedCorners:true}},
      catAxisLabelColor:'64748B',valAxisLabelColor:'64748B',
      valGridLine:{{color:'E2E8F0',size:0.5}},catGridLine:{{style:'none'}},
      showValue:true,dataLabelColor:'{pri}',dataLabelFontSize:10,
      showLegend:true,legendPos:'b',legendFontSize:10,
      showTitle:true,title:'월별 장애건수 비교 ({year_last}년 vs {year_this}년)',
      titleFontSize:11,titleColor:'{pri}'}});

  // 월별 추이 텍스트 요약
  sl.addShape('rect',{{x:6.4,y:3.12,w:3.2,h:2.18,fill:{{color:'FFFFFF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});
  sl.addShape('rect',{{x:6.4,y:3.12,w:3.2,h:0.36,fill:{{color:'{pri}'}},line:{{color:'{pri}'}}}});
  sl.addText('월별 추이 요약',{{x:6.5,y:3.13,w:3.0,h:0.34,fontSize:11,bold:true,color:'FFFFFF',fontFace:'Arial',margin:0}});
  sl.addText('{m_trend_str}',{{x:6.5,y:3.55,w:3.0,h:0.9,fontSize:10,color:'{pri}',fontFace:'Arial',margin:3,valign:'top'}});
  sl.addText('주의 역사 :',{{x:6.5,y:4.5,w:0.8,h:0.22,fontSize:9.5,bold:true,color:'{pri}',fontFace:'Arial',margin:0}});
  sl.addText('{warn_str}',{{x:7.3,y:4.5,w:2.2,h:0.22,fontSize:9.5,color:'C0392B',fontFace:'Arial',margin:0}});
  sl.addText('(전년 대비 증가 역사)',{{x:6.5,y:4.72,w:3.0,h:0.18,fontSize:8.5,color:'888888',fontFace:'Arial',margin:0}});
}}
"""

    # ── 슬라이드 B: 장치별 비교 + 상세 분석의견 ────────────────────────────
    # 분석의견을 별도 단락으로 구성 (addText 여러 번 호출)
    # 1) 전체 요약
    op1 = f"{year_this}년 {quarter} 총 {tt}건 (전년 {tl}건 대비 {td:+d}건, {tp_sign})"
    op2 = f"1호선 : {l1.get('this',0)}건 / 전년 {l1.get('last',0)}건 ({l1d:+d}건, {l1.get('pct',0.0):+.1f}%)"
    op3 = f"2호선 : {l2.get('this',0)}건 / 전년 {l2.get('last',0)}건 ({l2d:+d}건, {l2.get('pct',0.0):+.1f}%)"
    # 2) 장치별
    op4 = f"증가 장치 : {inc_str}"
    op5 = f"감소 장치 : {dec_str}"
    # 3) 주의 역사
    op6 = f"주의 역사 (전년 대비 증가) : {warn_str}"
    # 4) 종합 결론
    op7 = f"전년 대비 전체 {trend} 추세이나, {f_increased[0][0] if f_increased else '일부 장치'}는 증가하여 중점 관리 필요"

    slide_b = f"""
// ── 전년 동기 비교 (슬라이드 2/2): 장치별 비교 + 상세 분석의견
{{ const sl = pres.addSlide();
  sl.background = {{color:'{slide_bg}'}};
  {header_js}
  sl.addText('{year_last}년 {quarter} 대비 장치별 변화 및 상세 분석',{{x:0.55,y:0.62,w:8.5,h:0.28,fontSize:11,color:'64748B',fontFace:'Arial',margin:0}});
  ft(sl,'전년 비교 2/2');

  // 장치별 비교 가로 막대 차트 (좌측)
  sl.addChart(pres.charts.BAR,
    [{{name:'{year_this}년',labels:{_j(f_labels)},values:{_j(f_this)}}},
     {{name:'{year_last}년',labels:{_j(f_labels)},values:{_j(f_last)}}}],
    {{x:0.4,y:1.05,w:5.5,h:3.8,barDir:'bar',barGrouping:'clustered',
      chartColors:['{sec}','BBCCDD'],
      chartArea:{{fill:{{color:'FFFFFF'}},roundedCorners:true}},
      catAxisLabelColor:'64748B',valAxisLabelColor:'64748B',catAxisFontSize:10,
      valGridLine:{{color:'E2E8F0',size:0.5}},catGridLine:{{style:'none'}},
      showValue:true,dataLabelFontSize:10,dataLabelColor:'{pri}',
      showLegend:true,legendPos:'b',legendFontSize:10,
      showTitle:true,title:'장치별 장애건수 비교 상위 5종 ({year_last}년 vs {year_this}년)',
      titleFontSize:11,titleColor:'{pri}'}});

  // 상세 분석의견 (우측, 단락 분리)
  sl.addShape('rect',{{x:6.1,y:1.05,w:3.5,h:3.8,fill:{{color:'FFFFFF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});
  sl.addShape('rect',{{x:6.1,y:1.05,w:3.5,h:0.38,fill:{{color:'{pri}'}},line:{{color:'{pri}'}}}});
  sl.addText('상세 분석의견',{{x:6.2,y:1.06,w:3.3,h:0.36,fontSize:12,bold:true,color:'FFFFFF',fontFace:'Arial',margin:0}});

  // 단락 1: 전체 요약
  sl.addShape('rect',{{x:6.15,y:1.5,w:0.06,h:0.28,fill:{{color:'{acc}'}},line:{{color:'{acc}'}}}});
  sl.addText('전체 요약',{{x:6.25,y:1.5,w:3.3,h:0.25,fontSize:10,bold:true,color:'{pri}',fontFace:'Arial',margin:0}});
  sl.addText({_q(op1)},{{x:6.2,y:1.76,w:3.35,h:0.24,fontSize:9.5,color:'374151',fontFace:'Arial',margin:2}});
  sl.addText({_q(op2)},{{x:6.2,y:2.0,w:3.35,h:0.22,fontSize:9.5,color:'374151',fontFace:'Arial',margin:2}});
  sl.addText({_q(op3)},{{x:6.2,y:2.22,w:3.35,h:0.22,fontSize:9.5,color:'374151',fontFace:'Arial',margin:2}});

  // 구분선
  sl.addShape('rect',{{x:6.2,y:2.5,w:3.3,h:0.01,fill:{{color:'E2E8F0'}},line:{{color:'E2E8F0'}}}});

  // 단락 2: 장치별
  sl.addShape('rect',{{x:6.15,y:2.57,w:0.06,h:0.28,fill:{{color:'{acc}'}},line:{{color:'{acc}'}}}});
  sl.addText('장치별 변화',{{x:6.25,y:2.57,w:3.3,h:0.25,fontSize:10,bold:true,color:'{pri}',fontFace:'Arial',margin:0}});
  sl.addText({_q(op4)},{{x:6.2,y:2.83,w:3.35,h:0.22,fontSize:9.5,color:'C0392B',fontFace:'Arial',margin:2}});
  sl.addText({_q(op5)},{{x:6.2,y:3.05,w:3.35,h:0.22,fontSize:9.5,color:'27AE60',fontFace:'Arial',margin:2}});

  // 구분선
  sl.addShape('rect',{{x:6.2,y:3.32,w:3.3,h:0.01,fill:{{color:'E2E8F0'}},line:{{color:'E2E8F0'}}}});

  // 단락 3: 주의 역사 + 종합
  sl.addShape('rect',{{x:6.15,y:3.39,w:0.06,h:0.28,fill:{{color:'C0392B'}},line:{{color:'C0392B'}}}});
  sl.addText('주의 역사',{{x:6.25,y:3.39,w:3.3,h:0.25,fontSize:10,bold:true,color:'C0392B',fontFace:'Arial',margin:0}});
  sl.addText({_q(op6)},{{x:6.2,y:3.65,w:3.35,h:0.28,fontSize:9.5,color:'C0392B',fontFace:'Arial',margin:2}});

  // 구분선
  sl.addShape('rect',{{x:6.2,y:3.97,w:3.3,h:0.01,fill:{{color:'E2E8F0'}},line:{{color:'E2E8F0'}}}});
  sl.addText({_q(op7)},{{x:6.2,y:4.0,w:3.35,h:0.5,fontSize:9.5,bold:true,color:'{pri}',fontFace:'Arial',margin:3,valign:'top'}});

  // 하단 요약 바
  sl.addShape('rect',{{x:0.4,y:5.05,w:9.2,h:0.3,fill:{{color:'EEF2F7'}},line:{{color:'D4DCE8',pt:1}}}});
  sl.addText('장치별 증가 장치 집중 관리 및 주의 역사 현장 점검 강화 필요',{{x:0.6,y:5.09,w:9.0,h:0.22,fontSize:9.5,bold:true,color:'{pri}',fontFace:'Arial',margin:0,valign:'middle'}});
}}
"""
    return slide_a + slide_b


# ── 반복 장애 예방점검 슬라이드 ───────────────────────────────────────────────


def _repeat_slide(repeats: list, slide_bg, pri, acc, title_kr, repeat_num: int = 6) -> str:
    """
    반복 장애 예방점검 권고 슬라이드.
    - 10개 이하: 1장 (KPI + 테이블 + 의견 + 조치)
    - 11개 이상: 2장 (1장: KPI+테이블 상반부 / 2장: 테이블 하반부+의견+조치)
    """
    if not repeats:
        return ''

    urgent  = [r for r in repeats if r.get('level') == '즉시 점검']
    caution = [r for r in repeats if r.get('level') == '예방 점검 권고']
    total_r = len(repeats)

    # 분석의견 단락
    top3       = repeats[:3]
    top3_items = [f"{r['yeoksa']} {r['fault_type']} ({r['count']}건)" for r in top3]
    op1 = f"분기 내 동일 역사 · 장애종류 3회 이상 반복 장애 총 {total_r}건 식별"
    op2 = "주요 반복 장애 : " + ",  ".join(top3_items)
    op3 = f"즉시 점검 대상 {len(urgent)}건은 차기 정기점검 시 최우선 처리"
    op4 = "반복 발생 장치는 단순 장애처리로는 재발 방지 불가"
    op5 = "부품 교체 · 근본 원인 점검을 통한 예방 조치 필요"

    # 조치방향
    guide1 = "즉시 점검 : 차기 정기점검 시 최우선 부품 교체 및 근본 원인 제거"
    guide2 = "예방 점검 권고 : 정밀 진단 후 마모·노후 부품 교체 계획 수립"
    guide3 = "반복 횟수 증가 시 : 예방 점검 권고 → 즉시 점검으로 자동 상향"

    # 테이블 헤더 JS
    tbl_header = (
        f"[{{text:'역사',options:{{bold:true,color:'FFFFFF',fill:{{color:'{pri}'}}}}}},"
        f"{{text:'장애종류',options:{{bold:true,color:'FFFFFF',fill:{{color:'{pri}'}}}}}},"
        f"{{text:'발생건수',options:{{bold:true,color:'FFFFFF',fill:{{color:'{pri}'}}}}}},"
        f"{{text:'판정',options:{{bold:true,color:'FFFFFF',fill:{{color:'{pri}'}}}}}}]"
    )

    def make_row(r):
        bg = 'FFF5F5' if r.get('level') == '즉시 점검' else 'FFFBF0'
        return (
            f"[{{text:'{r['yeoksa']}',options:{{fill:{{color:'{bg}'}}}}}},"
            f"'{r['fault_type']}',"
            f"{{text:'{r['count']}건',options:{{bold:true,color:'{r['color']}'}}}},"
            f"{{text:'{r['level']}',options:{{bold:true,color:'{r['color']}'}}}}]"
        )

    # 슬라이드 공통 헤더 JS
    def slide_header(part_suffix=''):
        title_txt = f"{repeat_num}) 반복장애 예방점검 권고{part_suffix}"
        return f"""
  sl.addShape('rect',{{x:0.4,y:0.22,w:0.06,h:0.55,fill:{{color:'{acc}'}},line:{{color:'{acc}'}}}});
  sl.addText('{title_txt}',{{x:0.55,y:0.18,w:8.5,h:0.45,fontSize:22,bold:true,color:'{pri}',fontFace:'Arial',margin:0,valign:'middle'}});
  sl.addText('분기 내 동일 역사 · 장애종류 3회 이상 반복 발생 장치 집중 관리',{{x:0.55,y:0.62,w:8.5,h:0.28,fontSize:11,color:'64748B',fontFace:'Arial',margin:0}});
  sl.addShape('rect',{{x:0.4,y:0.9,w:9.2,h:0.02,fill:{{color:'D4DCE8'}},line:{{color:'D4DCE8'}}}});
  ft(sl,'예방점검 권고');"""

    # KPI 4개 공통
    kpi_block = f"""
  sl.addShape('rect',{{x:0.4,y:1.04,w:2.1,h:1.18,fill:{{color:'FFFFFF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});
  sl.addShape('rect',{{x:0.4,y:1.04,w:2.1,h:0.07,fill:{{color:'{pri}'}},line:{{color:'{pri}'}}}});
  sl.addText('반복 장애 감지',{{x:0.5,y:1.13,w:1.9,h:0.24,fontSize:10,color:'64748B',fontFace:'Arial',margin:0}});
  sl.addText('{total_r}건',{{x:0.5,y:1.35,w:1.9,h:0.48,fontSize:28,bold:true,color:'{pri}',fontFace:'Arial Black',align:'center',margin:0}});
  sl.addText('{title_kr} 합계',{{x:0.5,y:1.8,w:1.9,h:0.2,fontSize:9,color:'888888',align:'center',fontFace:'Arial',margin:0}});

  sl.addShape('rect',{{x:2.65,y:1.04,w:2.1,h:1.18,fill:{{color:'FFFFFF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});
  sl.addShape('rect',{{x:2.65,y:1.04,w:2.1,h:0.07,fill:{{color:'C0392B'}},line:{{color:'C0392B'}}}});
  sl.addText('즉시 점검',{{x:2.75,y:1.13,w:1.9,h:0.24,fontSize:10,color:'64748B',fontFace:'Arial',margin:0}});
  sl.addText('{len(urgent)}건',{{x:2.75,y:1.35,w:1.9,h:0.48,fontSize:28,bold:true,color:'C0392B',fontFace:'Arial Black',align:'center',margin:0}});
  sl.addText('5회 이상 반복',{{x:2.75,y:1.8,w:1.9,h:0.2,fontSize:9,color:'888888',align:'center',fontFace:'Arial',margin:0}});

  sl.addShape('rect',{{x:4.9,y:1.04,w:2.1,h:1.18,fill:{{color:'FFFFFF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});
  sl.addShape('rect',{{x:4.9,y:1.04,w:2.1,h:0.07,fill:{{color:'E57C0B'}},line:{{color:'E57C0B'}}}});
  sl.addText('예방 점검 권고',{{x:5.0,y:1.13,w:1.9,h:0.24,fontSize:10,color:'64748B',fontFace:'Arial',margin:0}});
  sl.addText('{len(caution)}건',{{x:5.0,y:1.35,w:1.9,h:0.48,fontSize:28,bold:true,color:'E57C0B',fontFace:'Arial Black',align:'center',margin:0}});
  sl.addText('3~4회 반복',{{x:5.0,y:1.8,w:1.9,h:0.2,fontSize:9,color:'888888',align:'center',fontFace:'Arial',margin:0}});

  sl.addShape('rect',{{x:7.15,y:1.04,w:2.45,h:1.18,fill:{{color:'EEF5FF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});
  sl.addShape('rect',{{x:7.15,y:1.04,w:2.45,h:0.07,fill:{{color:'2C5F8A'}},line:{{color:'2C5F8A'}}}});
  sl.addText('판정 기준',{{x:7.25,y:1.13,w:2.25,h:0.24,fontSize:10,color:'64748B',fontFace:'Arial',margin:0}});
  sl.addText('즉시 : 5회 이상',{{x:7.25,y:1.4,w:2.25,h:0.24,fontSize:10,bold:true,color:'C0392B',fontFace:'Arial',margin:0}});
  sl.addText('권고 : 3~4회',{{x:7.25,y:1.62,w:2.25,h:0.24,fontSize:10,bold:true,color:'E57C0B',fontFace:'Arial',margin:0}});
  sl.addText('(동일 역사 + 장애종류 기준)',{{x:7.25,y:1.85,w:2.25,h:0.18,fontSize:8,color:'888888',fontFace:'Arial',margin:0}});"""

    # 의견+조치 블록 (공통 재사용)
    def opinion_guide_block(table_end_y):
        # 의견박스: table_end_y 기준 오른쪽
        op_y  = 2.35
        op_h  = table_end_y - op_y - 0.06
        # 조치방향 y: table_end_y + 0.1 (푸터 y:5.35 침범 않도록 최대 4.72)
        guide_y = min(table_end_y + 0.1, 4.72)
        return f"""
  sl.addShape('rect',{{x:6.2,y:{op_y},w:3.4,h:{op_h:.2f},fill:{{color:'FFFFFF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});
  sl.addShape('rect',{{x:6.2,y:{op_y},w:3.4,h:0.34,fill:{{color:'{pri}'}},line:{{color:'{pri}'}}}});
  sl.addText('분석 의견',{{x:6.3,y:{op_y+0.01:.2f},w:3.2,h:0.32,fontSize:11,bold:true,color:'FFFFFF',fontFace:'Arial',margin:0}});
  sl.addText({_q(op1)},{{x:6.25,y:{op_y+0.4:.2f},w:3.3,h:0.26,fontSize:9.5,color:'{pri}',fontFace:'Arial',margin:2,bold:true}});
  sl.addText({_q(op2)},{{x:6.25,y:{op_y+0.66:.2f},w:3.3,h:0.26,fontSize:9.5,color:'374151',fontFace:'Arial',margin:2}});
  sl.addShape('rect',{{x:6.25,y:{op_y+0.96:.2f},w:3.3,h:0.01,fill:{{color:'E2E8F0'}},line:{{color:'E2E8F0'}}}});
  sl.addText({_q(op3)},{{x:6.25,y:{op_y+1.02:.2f},w:3.3,h:0.26,fontSize:9.5,color:'C0392B',fontFace:'Arial',margin:2,bold:true}});
  sl.addText({_q(op4)},{{x:6.25,y:{op_y+1.28:.2f},w:3.3,h:0.22,fontSize:9.5,color:'374151',fontFace:'Arial',margin:2}});
  sl.addText({_q(op5)},{{x:6.25,y:{op_y+1.5:.2f},w:3.3,h:0.22,fontSize:9.5,color:'374151',fontFace:'Arial',margin:2}});
  sl.addShape('rect',{{x:0.4,y:{guide_y:.2f},w:9.2,h:0.52,fill:{{color:'EEF2F7'}},line:{{color:'D4DCE8',pt:1}}}});
  sl.addShape('rect',{{x:0.4,y:{guide_y:.2f},w:0.06,h:0.52,fill:{{color:'{acc}'}},line:{{color:'{acc}'}}}});
  sl.addText('【조치 방향】',{{x:0.55,y:{guide_y+0.02:.2f},w:1.2,h:0.22,fontSize:9,bold:true,color:'{pri}',fontFace:'Arial',margin:0}});
  sl.addText({_q(guide1)},{{x:1.7,y:{guide_y+0.02:.2f},w:7.7,h:0.16,fontSize:8.5,color:'{pri}',fontFace:'Arial',margin:0}});
  sl.addText({_q(guide2)},{{x:1.7,y:{guide_y+0.17:.2f},w:7.7,h:0.16,fontSize:8.5,color:'{pri}',fontFace:'Arial',margin:0}});
  sl.addText({_q(guide3)},{{x:1.7,y:{guide_y+0.32:.2f},w:7.7,h:0.16,fontSize:8.5,color:'E57C0B',fontFace:'Arial',margin:0}});"""

    # ── 1장 처리 (10개 이하, 테이블이 슬라이드 안에 수용됨)
    # 1장 최대 수용 행: KPI 하단(y:2.35) ~ 조치방향 상단(y:4.72), 행높이 0.28
    # 가용높이: 4.72-2.35=2.37 → (2.37-0.34)/0.28 ≈ 7.25 → 7행
    MAX_SINGLE = 7   # 헤더 제외 최대 데이터 행

    if total_r <= MAX_SINGLE:
        # ── 1장 슬라이드: KPI + 테이블 + 의견 + 조치
        rows_parts = [tbl_header] + [make_row(r) for r in repeats]
        rows_js    = '[' + ',\n    '.join(rows_parts) + ']'
        tbl_h      = round(0.34 + len(repeats) * 0.28, 2)
        tbl_end_y  = round(2.35 + tbl_h, 2)

        return f"""
// ── 반복 장애 예방점검 권고 (1장)
{{ const sl = pres.addSlide();
  sl.background = {{color:'{slide_bg}'}};
  {slide_header()}
  {kpi_block}
  sl.addTable({rows_js},
    {{x:0.4,y:2.35,w:5.6,h:{tbl_h},border:{{pt:1,color:'D4DCE8'}},
      fill:{{color:'FFFFFF'}},colW:[1.2,1.6,0.9,1.9],
      fontFace:'Arial',fontSize:10.5,align:'center',rowH:0.28}});
  {opinion_guide_block(tbl_end_y)}
}}
"""
    else:
        # ── 2장 분리
        # 슬라이드 A: KPI + 전체 폭 테이블 (상위 9건, y:2.35~4.87 가용)
        # 테이블 가용: 4.87-2.35=2.52 → (2.52-0.34)/0.28 ≈ 7.8 → 최대 8행
        MAX_A = min(8, total_r)
        rows_a    = repeats[:MAX_A]
        parts_a   = [tbl_header] + [make_row(r) for r in rows_a]
        rows_js_a = '[' + ',\n    '.join(parts_a) + ']'
        tbl_h_a   = round(0.34 + len(rows_a) * 0.28, 2)

        # 슬라이드 B: 나머지 테이블 (헤더 없이) + 분석의견 + 조치
        # 테이블: y:1.05, 전체 폭 9.2 사용
        # 의견+조치: 테이블 아래에 가로 배치
        rows_b    = repeats[MAX_A:]
        parts_b   = [tbl_header] + [make_row(r) for r in rows_b]
        rows_js_b = '[' + ',\n    '.join(parts_b) + ']'
        tbl_h_b   = round(0.34 + len(rows_b) * 0.28, 2)
        tbl_end_b = round(1.05 + tbl_h_b + 0.1, 2)
        # 의견 + 조치를 테이블 오른쪽 또는 아래에 배치
        # 슬라이드 B는 헤더 공간(1.04) + 테이블 시작(1.05)이라 KPI 없음
        # → 테이블 좌측(w:5.5), 의견 우측(w:3.8) 나란히

        slide_a = f"""
// ── 반복 장애 예방점검 권고 (1/2) - KPI + 테이블 상위 {len(rows_a)}건
{{ const sl = pres.addSlide();
  sl.background = {{color:'{slide_bg}'}};
  {slide_header(' (1/2)')}
  {kpi_block}
  sl.addTable({rows_js_a},
    {{x:0.4,y:2.35,w:9.2,h:{tbl_h_a},border:{{pt:1,color:'D4DCE8'}},
      fill:{{color:'FFFFFF'}},colW:[2.2,2.8,1.5,2.7],
      fontFace:'Arial',fontSize:10.5,align:'center',rowH:0.28}});
  sl.addShape('rect',{{x:0.4,y:4.87,w:9.2,h:0.28,fill:{{color:'EEF2F7'}},line:{{color:'D4DCE8',pt:1}}}});
  sl.addText('※ 상위 {len(rows_a)}건 표시 — 다음 슬라이드에서 나머지 {len(rows_b)}건 및 분석의견 계속',
    {{x:0.6,y:4.9,w:9.0,h:0.22,fontSize:9.5,color:'{pri}',fontFace:'Arial',margin:0,bold:true}});
}}
"""
        # 슬라이드 B: 나머지 테이블(좌) + 의견(우) + 조치(하단)
        opinion_y  = 1.05
        opinion_h  = min(tbl_h_b, 3.55)   # 의견박스 높이 = 테이블 높이에 맞춤
        guide_y    = min(round(opinion_y + opinion_h + 0.1, 2), 4.72)

        slide_b = f"""
// ── 반복 장애 예방점검 권고 (2/2) - 나머지 {len(rows_b)}건 + 분석의견 + 조치방향
{{ const sl = pres.addSlide();
  sl.background = {{color:'{slide_bg}'}};
  {slide_header(' (2/2)')}
  sl.addTable({rows_js_b},
    {{x:0.4,y:{opinion_y},w:5.5,h:{tbl_h_b},border:{{pt:1,color:'D4DCE8'}},
      fill:{{color:'FFFFFF'}},colW:[1.2,1.6,0.9,1.8],
      fontFace:'Arial',fontSize:10.5,align:'center',rowH:0.28}});
  sl.addShape('rect',{{x:6.1,y:{opinion_y},w:3.5,h:{opinion_h:.2f},fill:{{color:'FFFFFF'}},line:{{color:'D4DCE8',pt:1}},shadow:ms()}});
  sl.addShape('rect',{{x:6.1,y:{opinion_y},w:3.5,h:0.34,fill:{{color:'{pri}'}},line:{{color:'{pri}'}}}});
  sl.addText('분석 의견',{{x:6.2,y:{opinion_y+0.01:.2f},w:3.3,h:0.32,fontSize:11,bold:true,color:'FFFFFF',fontFace:'Arial',margin:0}});
  sl.addText({_q(op1)},{{x:6.15,y:{opinion_y+0.4:.2f},w:3.4,h:0.26,fontSize:9.5,color:'{pri}',fontFace:'Arial',margin:2,bold:true}});
  sl.addText({_q(op2)},{{x:6.15,y:{opinion_y+0.66:.2f},w:3.4,h:0.28,fontSize:9.5,color:'374151',fontFace:'Arial',margin:2}});
  sl.addShape('rect',{{x:6.15,y:{opinion_y+0.98:.2f},w:3.4,h:0.01,fill:{{color:'E2E8F0'}},line:{{color:'E2E8F0'}}}});
  sl.addText({_q(op3)},{{x:6.15,y:{opinion_y+1.04:.2f},w:3.4,h:0.26,fontSize:9.5,color:'C0392B',fontFace:'Arial',margin:2,bold:true}});
  sl.addText({_q(op4)},{{x:6.15,y:{opinion_y+1.30:.2f},w:3.4,h:0.24,fontSize:9.5,color:'374151',fontFace:'Arial',margin:2}});
  sl.addText({_q(op5)},{{x:6.15,y:{opinion_y+1.54:.2f},w:3.4,h:0.24,fontSize:9.5,color:'374151',fontFace:'Arial',margin:2}});
  sl.addShape('rect',{{x:0.4,y:{guide_y:.2f},w:9.2,h:0.52,fill:{{color:'EEF2F7'}},line:{{color:'D4DCE8',pt:1}}}});
  sl.addShape('rect',{{x:0.4,y:{guide_y:.2f},w:0.06,h:0.52,fill:{{color:'{acc}'}},line:{{color:'{acc}'}}}});
  sl.addText('【조치 방향】',{{x:0.55,y:{guide_y+0.02:.2f},w:1.2,h:0.22,fontSize:9,bold:true,color:'{pri}',fontFace:'Arial',margin:0}});
  sl.addText({_q(guide1)},{{x:1.7,y:{guide_y+0.02:.2f},w:7.7,h:0.16,fontSize:8.5,color:'{pri}',fontFace:'Arial',margin:0}});
  sl.addText({_q(guide2)},{{x:1.7,y:{guide_y+0.17:.2f},w:7.7,h:0.16,fontSize:8.5,color:'{pri}',fontFace:'Arial',margin:0}});
  sl.addText({_q(guide3)},{{x:1.7,y:{guide_y+0.32:.2f},w:7.7,h:0.16,fontSize:8.5,color:'E57C0B',fontFace:'Arial',margin:0}});
}}
"""
        return slide_a + slide_b
