FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for better caching
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files (exclude unnecessary files for smaller image)
COPY *.py ./
COPY requirements.txt ./

# Create logs directory with proper permissions
RUN mkdir -p /app/logs && chmod 755 /app/logs

# Create non-root user for security
RUN useradd -m -u 1000 appuser && chown -R appuser:appuser /app
USER appuser

# Set environment variables
ENV PYTHONPATH=/app
ENV TZ=America/New_York
ENV PYTHONUNBUFFERED=1

# Health check - check if main modules can be imported
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import generate_report, config, attendance_tracker; print('Health check passed')" || exit 1

# Default command
CMD ["python", "main.py"] 