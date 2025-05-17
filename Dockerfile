# Use an official Python runtime as a parent image
FROM python:3.11-slim-buster

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file into the container
COPY requirements.txt .

# Install the Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the application code into the container
COPY . .

# Expose the port that Streamlit runs on (default is 8501, but we'll use 8080 for Cloud Run/App Engine)
EXPOSE 8080

# Command to run the Streamlit application
CMD ["streamlit", "run", "--server.port=8080", "--server.address=0.0.0.0", "exam_app.py"]