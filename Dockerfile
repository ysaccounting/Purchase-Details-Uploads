FROM python:3.11-slim

WORKDIR /app

# Install Python dependencies first (cached layer as long as requirements.txt is unchanged)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application
COPY . .

# Railway provides $PORT at runtime; default to 8080 for local runs
ENV PORT=8080
EXPOSE 8080

CMD gunicorn app:app --workers 1 --worker-class gthread --threads 4 --timeout 120 --bind 0.0.0.0:$PORT
