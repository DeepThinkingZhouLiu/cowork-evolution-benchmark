FROM python:3.11-slim-bookworm

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       libreoffice-calc \
       poppler-utils \
       fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir openpyxl==3.1.5 pillow==11.1.0

WORKDIR /app

COPY input_files/ /app/data/input_files/
RUN mkdir -p /app/output

RUN mkdir -p /data && ln -s /app/data/input_files /data/input_files
