# Use an official Python runtime as base image
FROM python:3.9-slim

# Set the working directory
WORKDIR /app

# Copy dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Expose port 8080 (important for Cloud Run)
EXPOSE 8080

# Start the Streamlit application
CMD ["python", "app.py"]
