FROM python:3.11-slim

ARG PIP_INDEX_URL
ARG PIP_DEFAULT_TIMEOUT=120

RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-calc poppler-utils \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir openpyxl==3.1.5

WORKDIR /app

COPY input_files/ /app/data/input_files/
RUN mkdir -p /app/output

RUN mkdir -p /data && ln -s /app/data/input_files /data/input_files
