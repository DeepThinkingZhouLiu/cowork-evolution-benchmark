FROM python:3.11-slim-bookworm

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-impress poppler-utils fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir python-pptx==1.0.2 Pillow==11.3.0 lxml==6.0.0

WORKDIR /app
COPY input_files /app/input_files
RUN mkdir -p /app/output /logs/artifacts/output

RUN mkdir -p /data && ln -s /app/input_files /data/input_files
