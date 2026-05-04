FROM python:3.9-slim

WORKDIR /app

# Install system dependencies if needed
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    python3-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Fix potential Windows line endings in startup script
RUN sed -i 's/\r$//' startup.sh && chmod +x startup.sh

# Railway uses the PORT environment variable
ENV PORT 8080
ENV PYTHONUNBUFFERED 1

# Use startup script to initialize database before starting app
# We use the shell form to ensure $PORT is expanded correctly
# We remove --preload to ensure gunicorn binds to the port as quickly as possible
CMD ./startup.sh gunicorn --bind 0.0.0.0:$PORT --workers 1 --worker-class gthread --threads 4 --timeout 120 app:app