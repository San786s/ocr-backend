FROM python:3.10-slim

# System dependencies
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    libglib2.0-0 \
    libsm6 \
    libxrender1 \
    libxext6 \
    && apt-get clean

# Optional: install additional language packs (e.g., Hindi, Tamil)
RUN apt-get install -y tesseract-ocr-hin tesseract-ocr-tam

# Set work directory
WORKDIR /code

# Copy code
COPY . /code

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose port
EXPOSE 7860

# Start the Flask app
CMD ["python", "app/main.py"]
