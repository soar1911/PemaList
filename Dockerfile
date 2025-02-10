FROM python:3.9-slim

WORKDIR /app

# 安裝系統依賴
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 創建必要的目錄
RUN mkdir -p /app/output_files

COPY . .

# 設置權限
RUN chown -R root:root /app && \
    chmod -R 755 /app && \
    chmod -R 777 /app/output_files

EXPOSE 33080

CMD ["gunicorn", "--bind", "0.0.0.0:33080", "--log-level", "debug", "--timeout", "120", "app:app"]
