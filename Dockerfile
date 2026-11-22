# Production Dockerfile for doc-converter
FROM python:3.11-slim

# Prevent Python from writing .pyc files and buffer output
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8000

# Install LibreOffice headless, curl for health checks, and standard fonts for visual fidelity
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-liberation \
    fonts-dejavu-core \
    curl \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy dependency specifications first to leverage Docker layer caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy all application code
COPY . .

# Expose default port
EXPOSE 8000

# Container healthcheck
HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 \
    CMD curl -f http://localhost:${PORT:-8000}/health || exit 1

# Start Uvicorn bound to 0.0.0.0:${PORT:-8000}
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}"]
