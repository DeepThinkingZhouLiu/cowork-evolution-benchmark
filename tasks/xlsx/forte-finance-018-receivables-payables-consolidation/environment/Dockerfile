FROM python:3.11-slim-bookworm

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       libreoffice-calc \
       poppler-utils \
       fonts-noto-cjk \
       fonts-liberation2 \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir openpyxl==3.1.5 pypinyin==0.55.0

WORKDIR /app
