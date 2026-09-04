FROM python:3.11-slim
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-calc poppler-utils \
    && rm -rf /var/lib/apt/lists/*
RUN pip install --no-cache-dir openpyxl==3.1.5 python-docx==1.1.2 pillow==10.4.0
RUN mkdir -p /app/output /data
COPY input_files /app/data/input_files
RUN ln -s /app/data/input_files /data/input_files
WORKDIR /app
