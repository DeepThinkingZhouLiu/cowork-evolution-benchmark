FROM python:3.11.9-slim-bookworm

RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-calc poppler-utils fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/* \
    && pip install --no-cache-dir openpyxl==3.1.5

WORKDIR /app

COPY input_files/ /app/data/input_files/
RUN mkdir -p /app/output
RUN mkdir -p /app/data/input_files /app/output

RUN mkdir -p /data && ln -s /app/data/input_files /data/input_files
