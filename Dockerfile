FROM python:3.11-slim

WORKDIR /app

# 安装必要工具
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        ca-certificates \
        && \
    rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

# 创建非 root 用户运行
RUN useradd -m -u 1000 appuser && chown -R appuser:appuser /app
USER appuser

EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]