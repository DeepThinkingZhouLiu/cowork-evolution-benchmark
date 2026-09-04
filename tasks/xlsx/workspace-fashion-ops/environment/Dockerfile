FROM python:3.11-slim
RUN pip install --no-cache-dir openpyxl==3.1.5
WORKDIR /app
COPY input_files /app/data/input_files
COPY input_files /data/input_files
RUN mkdir -p /app/output
