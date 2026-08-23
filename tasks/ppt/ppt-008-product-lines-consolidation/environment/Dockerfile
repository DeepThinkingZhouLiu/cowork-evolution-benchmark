FROM python:3.11-slim-bookworm

WORKDIR /app
RUN pip install --no-cache-dir \
    python-pptx==1.0.2 \
    openpyxl==3.1.5 \
    python-docx==1.2.0 \
    Pillow==11.3.0 \
    PyMuPDF==1.26.4
RUN mkdir -p /app/output /app/data/input_files
COPY data/input_files/ /app/data/input_files/
