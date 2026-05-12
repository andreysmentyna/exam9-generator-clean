FROM python:3.11-slim-bookworm

# Устанавливаем LibreOffice 24.2 из официальных Debian Backports
RUN echo "deb http://deb.debian.org/debian bookworm-backports main" >> /etc/apt/sources.list && \
    apt-get update && \
    apt-get install -y --no-install-recommends -t bookworm-backports \
        libreoffice-writer \
        libreoffice-math \
    && \
    # Устанавливаем шрифты для кириллицы и математических символов
    apt-get install -y --no-install-recommends \
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

CMD ["gunicorn", "app:app", "-b", "0.0.0.0:10000", "--timeout", "600"]