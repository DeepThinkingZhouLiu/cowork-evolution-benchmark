FROM python:3.12-slim

# 正式任务必须固定全部依赖版本，并安装生成与检查办公文件所需的工具。
WORKDIR /app

COPY input_files/ /app/data/input_files/
RUN mkdir -p /app/output
RUN mkdir -p /app/output /app/data/input_files

RUN mkdir -p /data && ln -s /app/data/input_files /data/input_files
