import requests

url = "http://localhost:8080/cadcorr/findnome"
nome = "CRISTIANO MARIANO DA SILVA"
querystring = {"nome":{nome}}

payload = ""
headers = {"User-Agent": "insomnia/2023.5.8"}

response = requests.request("GET", url, data=payload, headers=headers, params=querystring)

print(response.text)