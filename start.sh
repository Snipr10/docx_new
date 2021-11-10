#!/bin/bash
docker build -t flask-docx .

docker run -d -p 5000:5000 flask-docx

