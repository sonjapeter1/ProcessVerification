FROM python:3.12-slim

WORKDIR /app
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Copy wheels first
COPY wheels /wheels

# Copy requirements
COPY requirements.txt .

# Install from local wheels only
RUN ls -l /wheels && pip install --no-index --find-links=/wheels -r requirements.txt

# Copy application code
COPY . .

CMD ["gunicorn", "-w", "2", "-k", "gthread", "--threads", "8", "-b", "0.0.0.0:8000", "--access-logfile", "-", "--error-logfile", "-", "wsgi:app"]
