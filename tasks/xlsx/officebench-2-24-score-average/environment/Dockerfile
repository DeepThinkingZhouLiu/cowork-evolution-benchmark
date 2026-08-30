FROM python:3.11-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
RUN pip install --no-cache-dir openpyxl==3.1.5
WORKDIR /app
COPY input_files/ /app/data/input_files/
RUN mkdir -p /app/output /data && ln -s /app/data/input_files /data/input_files
