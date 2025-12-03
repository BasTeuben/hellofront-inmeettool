# ================================
# Force rebuild (cache busting)
# ================================
# rebuild-force-001

FROM python:3.11-slim

# ================================
# Maak werkdirectory
# ================================
WORKDIR /app

# ================================
# Kopieer ALLE bestanden expliciet
# (dit voorkomt Railway/Nixpacks context-bugs)
# ================================
COPY ./ /app/

# ================================
# Installeer dependencies
# ================================
RUN pip install --upgrade pip
RUN pip install -r requirements.txt

# ================================
# Expose Streamlit port
# ================================
EXPOSE 8501

# ================================
# Start Streamlit expliciet
# ================================
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
