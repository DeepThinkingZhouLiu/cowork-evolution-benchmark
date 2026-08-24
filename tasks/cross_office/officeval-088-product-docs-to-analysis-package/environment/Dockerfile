FROM python:3.11-slim-bookworm

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       libreoffice-writer \
       libreoffice-calc \
       poppler-utils \
       tesseract-ocr \
       tesseract-ocr-chi-sim \
       fonts-liberation2 \
       fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir \
    openpyxl==3.1.5 \
    python-docx==1.1.2 \
    Pillow==11.1.0

WORKDIR /app
