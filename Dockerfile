FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    ALL4CAR_HEADLESS=0 \
    ALL4CAR_FAST_MODE=1 \
    DISPLAY=:99 \
    ALL4CAR_DATA_DIR=/var/data/all4car_batch_scraper

RUN apt-get update && apt-get install -y --no-install-recommends \
    chromium \
    chromium-driver \
    xvfb \
    ca-certificates \
    fonts-liberation \
    fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

COPY . .
RUN mkdir -p /var/data/all4car_batch_scraper

EXPOSE 8501

CMD ["sh", "-c", "Xvfb :99 -screen 0 1920x1080x24 -ac +extension RANDR >/tmp/xvfb.log 2>&1 & sleep 1 && exec streamlit run app.py --server.address=0.0.0.0 --server.port=${PORT:-8501} --server.headless=true --browser.gatherUsageStats=false"]
