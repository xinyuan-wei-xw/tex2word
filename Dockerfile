FROM python:3.12-slim

# system deps: pandoc + tesseract (eng + chi_sim) for the OCR mode
RUN apt-get update && apt-get install -y --no-install-recommends \
    pandoc \
    tesseract-ocr tesseract-ocr-eng tesseract-ocr-chi-sim \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

# run as non-root `user` (Hugging Face Docker Spaces convention)
RUN useradd -m -u 1000 user && chown -R user:user /app
USER user

# HF Spaces / generic hosts provide $PORT; default 7860
ENV PORT=7860
EXPOSE 7860
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-7860}"]
