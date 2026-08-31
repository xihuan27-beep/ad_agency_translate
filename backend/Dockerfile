# Build context is the repository root (see render.yaml) so this Dockerfile
# can reach both backend/ and the shared ai_utils.py / pptx_utils.py / docx_utils.py.
FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils \
    fonts-noto-cjk \
    fonts-nanum \
    fonts-unfonts-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY backend/requirements.txt ./backend/requirements.txt
RUN pip install --no-cache-dir -r backend/requirements.txt

COPY ai_utils.py pptx_utils.py docx_utils.py ./
COPY backend/ ./backend/

EXPOSE 8000
CMD ["uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "8000"]
