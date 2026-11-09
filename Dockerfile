FROM python:3.12-slim

# LibreOffice sin interfaz + fuentes compatibles con las de Office
# (Carlito = Calibri, Caladea = Cambria, Liberation = Arial/Times/Courier).
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice-writer libreoffice-calc libreoffice-impress libreoffice-draw \
      fonts-dejavu fonts-liberation fonts-liberation2 fonts-crosextra-carlito fonts-crosextra-caladea \
      fonts-noto-core fonts-noto-color-emoji \
    && rm -rf /var/lib/apt/lists/*

# Hugging Face Spaces corre el contenedor con el usuario 1000.
RUN useradd -m -u 1000 app
WORKDIR /home/app/srv
COPY --chown=app requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY --chown=app app.py .

USER app
ENV HOME=/home/app PORT=7860
EXPOSE 7860
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT} --proxy-headers"]
