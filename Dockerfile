FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PORT=8080

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8080

# Hosts such as Render inject their own $PORT; 8080 is the default. One worker keeps in-memory job state consistent.
# --proxy-headers lets the app see it is being served over HTTPS.
CMD exec uvicorn main:app --host 0.0.0.0 --port ${PORT:-8080} --workers 1 \
    --proxy-headers --forwarded-allow-ips "*"
