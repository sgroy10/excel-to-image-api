FROM python:3.11-slim

# Install LibreOffice, poppler-utils, and fonts
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    fonts-liberation \
    fonts-dejavu \
    fontconfig \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    && fc-cache -f -v

# Create a non-root user for LibreOffice
RUN useradd -m -s /bin/bash appuser

# Set working directory
WORKDIR /app

# Copy requirements and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY main.py .

# Create temp directory with proper permissions
RUN mkdir -p /tmp/office && chown -R appuser:appuser /tmp/office

# Set environment variables for LibreOffice
ENV HOME=/tmp/office
ENV TMPDIR=/tmp/office

# Change to non-root user
USER appuser

# Expose port
EXPOSE 8000

# Run the application
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
