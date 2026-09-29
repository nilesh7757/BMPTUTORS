FROM python:3.11-slim

WORKDIR /app

# Install system utilities
RUN apt-get update && apt-get install -y --no-install-recommends \
    sqlite3 gzip \
    && rm -rf /var/lib/apt/lists/*

# Install python requirements
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code and assets
COPY app/ ./app/
COPY cleaned_tutors_data.db.gz .
RUN gzip -d -k cleaned_tutors_data.db.gz
COPY BMP_Tutors_Cleaned_Master.xlsx .
COPY BMP_Tutors_Verified_Active_Master.xlsx .
COPY BMP_Tutors_Visual_Analytics_Report.docx .
COPY BMP_Tutors_Visual_Analytics_Complete_Report.pdf .
COPY extracted_images/ ./extracted_images/

# Default port for Hugging Face Spaces is 7860; Render / Koyeb uses $PORT
ENV PORT=7860
ENV PYTHONPATH="/app:/app/app:${PYTHONPATH}"
EXPOSE 7860

# Launch server
CMD ["sh", "-c", "python3 -m uvicorn app.server:app --host 0.0.0.0 --port ${PORT:-7860}"]
