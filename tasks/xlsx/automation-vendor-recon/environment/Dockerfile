FROM python:3.12-slim
RUN pip install --no-cache-dir openpyxl==3.1.5
RUN mkdir -p /app/output /app/data/input_files /data/input_files
COPY input_files/ /app/data/input_files/
COPY input_files/ /data/input_files/
WORKDIR /app
