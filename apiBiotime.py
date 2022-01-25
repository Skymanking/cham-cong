import json
import requests


url = "http://chamcong.pte.vn:8888/iclock/api/transactions/?start_time=2022-01-10"
# use General token
headers = {
    "Content-Type": "application/json",
    "Authorization": "token 9ea1dfbad18d482738bedf550c8539cff2a3a2fb",
}
data = {

    "start_time": "2022-01-10",
    "end_time": "2022-01-20",
}


response = requests.get(url, data=json.dumps(data), headers=headers)
print(response.text)