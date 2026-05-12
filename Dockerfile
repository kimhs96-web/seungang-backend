FROM python:3.11-slim

# Node.js 설치
RUN apt-get update && apt-get install -y curl && \
    curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
    apt-get install -y nodejs && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Python 패키지
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Node 패키지 (pptxgenjs)
COPY package.json .
RUN npm install

# 앱 소스
COPY . .

EXPOSE 8000
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "1"]
