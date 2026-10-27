# Debian's python3-uno package is built for the system Python, so use Debian Python
# and a venv that can see system site-packages.
FROM python:3.11-slim-bookworm

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PATH="/opt/venv/bin:${PATH}"

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
      python3 \
      python3-venv \
      python3-uno \
      libreoffice-writer-nogui \
      pandoc \
      fonts-liberation \
    && rm -rf /var/lib/apt/lists/* \
    && /usr/bin/python3 -m venv --system-site-packages /opt/venv


COPY requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY docx_to_md_api.py ./docx_to_md_api.py

# Run as an unprivileged user. LibreOffice needs a writable HOME/profile.
RUN useradd --create-home --uid 10001 appuser \
    && mkdir -p /tmp/docx-convert /app/output \
    && chown -R appuser:appuser /app /home/appuser /tmp/docx-convert
ENV HOME=/home/appuser OUTPUT_DIR=/app/output
USER appuser

EXPOSE 8000
CMD ["uvicorn", "docx_to_md_api:app", "--host", "0.0.0.0", "--port", "8000"]
