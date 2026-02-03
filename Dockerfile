FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Make startup script executable
RUN chmod +x startup.sh

EXPOSE 8000

# Use startup script to initialize database before starting app
CMD ["sh", "-c", "./startup.sh && exec gunicorn --bind 0.0.0.0:8000 --workers 2 app:app"]