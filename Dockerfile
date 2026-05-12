FROM python:3.11-slim-bookworm

RUN apt-get update && apt-get install -y --no-install-recommends gnupg wget && \
    echo "deb http://download.documentfoundation.org/libreoffice/stable/24.2/deb/ bookworm main" > /etc/apt/sources.list.d/libreoffice.list && \
    wget -qO - https://download.documentfoundation.org/libreoffice/stable/24.2/deb/Release.key | apt-key add - && \
    apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice-writer \
        libreoffice-math \
        fonts-liberation \
        fonts-symbola \
        fontconfig && \
    fc-cache -fv && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PORT 10000
EXPOSE 10000

CMD ["gunicorn", "app:app", "-b", "0.0.0.0:10000"]