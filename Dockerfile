FROM python:3.12-slim

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    TZ=Asia/Tokyo \
    HOME=/tmp

RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice-calc \
      libreoffice-core \
      fonts-noto-cjk \
      fontconfig \
      tzdata \
      curl \
 && rm -rf /var/lib/apt/lists/*

COPY 61-jp-priority.conf /etc/fonts/conf.d/61-jp-priority.conf
RUN fc-cache -f

WORKDIR /srv
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY config.py excel_editor.py pdf_maker.py main.py ./

EXPOSE 8080
HEALTHCHECK --interval=60s --timeout=10s --start-period=180s --retries=3 \
  CMD curl -fsS http://127.0.0.1:8080/health || exit 1

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8080", "--workers", "1"]
