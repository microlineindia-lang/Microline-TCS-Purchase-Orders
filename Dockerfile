FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Cloud Run provides $PORT. One worker keeps job state consistent.
CMD exec uvicorn main:app --host 0.0.0.0 --port ${PORT:-8080} --workers 1
