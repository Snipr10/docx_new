# Docker 

docker build -t flask-docx .

docker run -d -p 5000:5000 flask-docx


# Example
http://localhost:5000/get_report?period=month&reference_ids=[370,%20369,%201056]