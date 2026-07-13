FROM python:3.11-slim
WORKDIR /app
# weasyprint 시스템 라이브러리 + 한글폰트 + OCR PDF 직독(poppler)
RUN apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils poppler-data \
    libpango-1.0-0 libpangocairo-1.0-0 libpangoft2-1.0-0 \
    libgdk-pixbuf-2.0-0 libffi-dev libcairo2 fonts-nanum \
    && rm -rf /var/lib/apt/lists/*
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
CMD ["sh","-c","uvicorn main:app --host 0.0.0.0 --port ${PORT:-8080}"]
