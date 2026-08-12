FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 DEBIAN_FRONTEND=noninteractive
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress libreoffice-core poppler-utils fonts-noto-cjk fonts-liberation \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py .
EXPOSE 10000
CMD ["sh","-c","streamlit run app.py --server.address=0.0.0.0 --server.port=${PORT:-10000} --server.headless=true --browser.gatherUsageStats=false"]
