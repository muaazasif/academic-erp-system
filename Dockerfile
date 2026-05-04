FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Make startup script executable
RUN chmod +x startup.sh

EXPOSE 8080

# Use startup script to initialize database before starting app
# Switch to gthread for better performance on memory-limited environments
# Use 1 worker to ensure memory isn't the issue during heavy initialization
CMD ["sh", "-c", "./startup.sh gunicorn --bind 0.0.0.0:${PORT} --workers 1 --worker-class gthread --threads 4 --timeout 120 --preload app:app"]