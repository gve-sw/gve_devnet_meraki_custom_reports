FROM python:3.9-slim-buster
RUN mkdir /app
WORKDIR /app
COPY ./requirements.txt /app
RUN pip install -r requirements.txt
COPY ./main /app
CMD ["python", "./main.py"]
