FROM python:3.8-slim-buster
# set work directory
COPY . /app
WORKDIR /app
RUN apt-get update && \
    apt-get upgrade -y && \
    apt-get install -y git
RUN pip3 install -r requirements.txt
#CMD [ "python3", "-m" , "flask", "run", "--host=0.0.0.0"]
EXPOSE 8080
ENTRYPOINT ["python"]
CMD ["app.py"]
