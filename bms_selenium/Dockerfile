FROM selenium/standalone-chrome:latest

USER root

# Install Python3 + pip only (Chrome + ChromeDriver already in base image)
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 \
    python3-pip \
    && rm -rf /var/lib/apt/lists/*

# Paths provided by selenium/standalone-chrome
ENV CHROME_BIN=/usr/bin/google-chrome
ENV CHROMEDRIVER_PATH=/usr/bin/chromedriver

WORKDIR /app

COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt --break-system-packages

COPY . .

CMD ["python3", "kirtan-baldev-meghna-report.py"]