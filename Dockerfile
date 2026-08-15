FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# One worker: LibreOffice conversion is the memory hog and 512MB does not fit two.
# 300s timeout: a cold LibreOffice start plus a multi-image report beats the 30s default.
CMD gunicorn app:app --bind 0.0.0.0:${PORT:-5000} --workers 1 --threads 2 --timeout 300

