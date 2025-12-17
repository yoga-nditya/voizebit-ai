FROM python:3.11-slim

# Install LibreOffice + font (WAJIB untuk PDF)
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      libreoffice \
      libreoffice-writer \
      fonts-dejavu \
      fonts-liberation && \
    rm -rf /var/lib/apt/lists/*

# Workdir
WORKDIR /app

# Install Python deps
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app
COPY . .

# Env untuk LibreOffice headless
ENV FLASK_ENV=production
ENV HOME=/tmp
ENV TMPDIR=/tmp

# Render inject PORT otomatis
CMD ["sh", "-c", "gunicorn -b 0.0.0.0:${PORT:-10000} app:app"]
