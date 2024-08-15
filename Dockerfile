FROM python:3.10-slim

# Enable i386 architecture and install dependencies
RUN dpkg --add-architecture i386 && apt-get update && apt-get install -y \
    wine32 \
    wget \
    xvfb \
    && rm -rf /var/lib/apt/lists/*

# Install winetricks from the official source
RUN wget https://raw.githubusercontent.com/Winetricks/winetricks/master/src/winetricks && \
    chmod +x winetricks && \
    mv winetricks /usr/local/bin/

# Install corefonts using winetricks (optional, but can help with some Wine issues)
RUN /usr/local/bin/winetricks -q corefonts

# Download the Python installer for Windows
RUN wget https://www.python.org/ftp/python/3.10.0/python-3.10.0-amd64.exe

# Run the Python installer with Wine using Xvfb for headless operation
RUN Xvfb :1 -screen 0 1024x768x24 & \
    DISPLAY=:1 WINEDEBUG=+all wine python-3.10.0-amd64.exe /quiet InstallAllUsers=1 PrependPath=1 2>&1 | tee /var/log/wine_install.log

# Set the WINEPATH environment variable to the installed Python directory
ENV WINEPATH C:\\users\\root\\AppData\\Local\\Programs\\Python\\Python310

# Use wine to install Python packages with pip
RUN DISPLAY=:1 wine ${WINEPATH}\\python.exe -m pip install pyinstaller xlwings PyQt5

# Set the working directory
WORKDIR /app

# Copy your application files into the container
COPY . /app

# Command to build your application using PyInstaller
CMD ["wine", "C:\\users\\root\\AppData\\Local\\Programs\\Python\\Python310\\Scripts\\pyinstaller.exe", "--onefile", "excel.py"]
