FROM python:3.12-slim
RUN pip install --no-cache-dir openpyxl==3.1.5
COPY input_files/ /app/data/input_files/
RUN mkdir -p /app/output /data/input_files && cp -a /app/data/input_files/. /data/input_files/
WORKDIR /app
