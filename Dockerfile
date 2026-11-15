FROM python:3.10-slim

WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

COPY app/ ./app/

RUN mkdir -p data/exports data/seller_exports data/seller_inbox data/seller_processed

ENV PYTHONUNBUFFERED=1

CMD ["python", "-m", "app.main"]
