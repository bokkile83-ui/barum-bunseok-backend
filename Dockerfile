FROM python:3.11-slim
WORKDIR /app
# weasyprint 시스템 라이브러리 + 한글폰트 + OCR PDF 직독(poppler)
RUN apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils poppler-data qpdf \
    libpango-1.0-0 libpangocairo-1.0-0 libpangoft2-1.0-0 \
    libgdk-pixbuf-2.0-0 libffi-dev libcairo2 fonts-nanum \
    && rm -rf /var/lib/apt/lists/*
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ★★★★★v463 제71조 (2026.08.17 실사고) — 「Deployment successful」인데 파일이 옛것이었다.
#   Build Logs 실측: `COPY . .  cached  0ms` — 도커가 소스 변경을 못 알아채고
#   옛 레이어를 그대로 썼다. 그래서 서버는 v443, GitHub은 v462였다.
#   → 이 값을 바꾸면 이 줄 아래 레이어의 캐시가 전부 무효화된다.
#     소스를 올려도 서버가 안 바뀌면 이 숫자를 하나 올린다.
ARG CACHEBUST=v474
RUN echo "cachebust=$CACHEBUST"

COPY . .
CMD ["sh","-c","uvicorn main:app --host 0.0.0.0 --port ${PORT:-8080}"]
