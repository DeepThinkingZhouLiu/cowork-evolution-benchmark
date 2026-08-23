FROM python:3.12-slim

WORKDIR /app
RUN python -m pip install --no-cache-dir openpyxl==3.1.5 python-docx==1.2.0
RUN mkdir -p /app/data/input_files /app/output
COPY data/input_files/ /app/data/input_files/
