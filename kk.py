import json
import requests


url = "http://chamcong.pte.vn:8888/api-token-auth/"
headers = {
    "Content-Type": "application/json",
}
data = {
    "username": "admin",
    "password": "admin123"
}

response = requests.post(url, data=json.dumps(data), headers=headers)
print(response.text)

