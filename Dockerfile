# Use official Python image
FROM python:3.11-slim

# Install system deps (optional: for Excel libraries that need libxml, etc.)
RUN apt-get update && apt-get install -y \
    build-essential \
    libxml2-dev \
    libxslt-dev \
    git \
    && rm -rf /var/lib/apt/lists/*

# Set work directory
WORKDIR /app

# Clone your repo (replace with your Git URL)
# Or you can COPY . if you’re building locally
RUN git clone https://github.com/your-username/your-repo.git repo

# Install dependencies (adjust based on your scripts)
COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Default command: just drop into bash
CMD ["/bin/bash"]
