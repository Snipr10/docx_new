FROM python:3.8-slim-buster
# set work directory
COPY . /app
WORKDIR /app
RUN apt-get update && \
    apt-get upgrade -y && \
    apt-get install -y git
RUN pip3 install -r requirements.txt
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8080"]
