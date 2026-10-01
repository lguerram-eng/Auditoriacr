# Muenra Vouching — API y trabajador de procesamiento
FROM python:3.11-slim AS base

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# Tesseract OCR (español + inglés) y fuentes para los datos de demostración
RUN apt-get update \
 && apt-get install -y --no-install-recommends tesseract-ocr tesseract-ocr-spa tesseract-ocr-eng fonts-dejavu-core curl \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt requirements-dev.txt ./
RUN pip install -r requirements-dev.txt

COPY . .

# Usuario sin privilegios
RUN useradd --create-home --uid 10001 muenra && mkdir -p /data/storage && chown -R muenra:muenra /data /app
USER muenra

ENV MUENRA_STORAGE_DIR=/data/storage \
    MUENRA_FRONTEND_DIR=/app/frontend-no-servido \
    MUENRA_INLINE_WORKER=false

EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --retries=5 CMD curl -fs http://localhost:8000/api/health || exit 1

# Aplica migraciones y arranca la API
CMD ["sh", "-c", "alembic upgrade head && uvicorn app.main:app --host 0.0.0.0 --port 8000 --proxy-headers --forwarded-allow-ips='*' --workers 2"]
