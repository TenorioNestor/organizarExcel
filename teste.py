import requests
import json

url = "http://localhost:8080/cadcorr/findnome"

querystring = {"nome":"APARECIDA BIAJOLI","cpf":"55505006834","email":"ros94@gmail.com"}

payload = ""
headers = {"User-Agent": "insomnia/2023.5.8"}

response = requests.request("GET", url, data=payload, headers=headers, params=querystring)

print(response.text)
data = response


# Decode the response data
data_string = data.text

# Parse the JSON string into a dictionary
data_dict = json.loads(data_string)

# Check if data_dict is a list
if isinstance(data_dict, list):
    # Iterate through the list and search for the dictionary with "nome"
    for item in data_dict:
        if "nome" in item:
            specific_line = item["nome"]
            print(specific_line)
            break  # Exit the loop after finding the first occurrence
else:
    print("Unexpected data structure. Expected a list of dictionaries.")