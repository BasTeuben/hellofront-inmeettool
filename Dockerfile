FROM python:3.11-slim

WORKDIR /app

COPY . .

RUN pip install --upgrade pip
RUN pip install -r requirements.txt

EXPOSE 8501

CMD ["streamlit", "run", "inmeetverwerker_hellofront.py", "--server.port=8501", "--server.address=0.0.0.0"]
