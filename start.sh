#!/bin/bash

docker build -t fast-api .

docker run -d --name fast-api-container -p 5000:5000 fast-api