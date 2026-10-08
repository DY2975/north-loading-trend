# 区域工厂液体装货趋势 —— 通用部署镜像
# 适用平台: Render / Railway / Fly.io / Koyeb / Hugging Face Spaces 等支持容器或 Python 的平台
FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    TZ=Asia/Shanghai

WORKDIR /app

# 依赖单独一层, 便于缓存
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 平台会注入自己的 PORT; 这里只是本地兜底
ENV PORT=8787
EXPOSE 8787

CMD ["python", "server.py"]
