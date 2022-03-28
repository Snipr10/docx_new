# Docker 

docker build -t fast-api .

docker run -d --name fast-api-container -p 5000:5000 fast-api

# Example
http://localhost:8080/get_report?period=week&reference_ids[]=370&reference_ids[]=369&reference_ids[]=1056&login=Kom_pvsmi@gov.spb.ru&password=m0fpoemZc3K1&thread_id=4188