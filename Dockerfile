FROM python:3.10-slim

# Устанавливаем шрифты Liberation (кириллица)
RUN apt-get update && \
    apt-get install -y --no-install-recommends fonts-liberation fontconfig && \
    fc-cache -fv && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT 10000
EXPOSE 10000

CMD ["gunicorn", "app:app", "-b", "0.0.0.0:10000"]