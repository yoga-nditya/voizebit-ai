# Dockerfile (with LibreOffice for docx->pdf conversion)
FROM python:3.11-slim

# set noninteractive to avoid tzdata prompts
ENV DEBIAN_FRONTEND=noninteractive
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PORT=5000

WORKDIR /app

# system deps for LibreOffice, unzip/zip and fonts
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      build-essential \
      unzip \
      zip \
      libreoffice-core \
      libreoffice-writer \
      libreoffice-common \
      libreoffice-java-common \
      fonts-dejavu-core \
      ca-certificates \
      wget \
      && apt-get clean && rm -rf /var/lib/apt/lists/*

# copy requirements first for better caching
COPY requirements.txt /app/requirements.txt
RUN pip install --upgrade pip && \
    pip install --no-cache-dir -r /app/requirements.txt

# copy app
COPY . /app

# create dirs (in case not present)
RUN mkdir -p /app/static/files /app/temp

# expose port
EXPOSE ${PORT}

# Use gunicorn for production; fallback to flask's dev server if desired.
# Ensure your app entrypoint is app:app (as in provided code).
CMD ["gunicorn", "--workers", "3", "--bind", "0.0.0.0:5000", "app:app"]
