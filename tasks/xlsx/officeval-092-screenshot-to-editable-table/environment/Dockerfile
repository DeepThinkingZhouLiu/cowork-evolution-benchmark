FROM python:3.11.9-slim-bookworm

RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-calc poppler-utils fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/* \
    && pip install --no-cache-dir openpyxl==3.1.5

WORKDIR /app
RUN mkdir -p /app/data/input_files /app/output
