FROM python:3.11-slim-bookworm

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       libreoffice-calc \
       libreoffice-impress \
       poppler-utils \
       fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir lxml==6.0.1 openpyxl==3.1.5 python-pptx==1.0.2

WORKDIR /app
