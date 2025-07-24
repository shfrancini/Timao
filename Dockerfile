FROM n8nio/n8n

USER root

# Install Python + pip + OpenPyXL
RUN apk add --no-cache python3 py3-pip

# Use --break-system-packages to bypass PEP 668 protection
RUN pip3 install --break-system-packages openpyxl python-docx

USER node
