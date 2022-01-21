import json
import requests


url = "http://chamcong.pte.vn:8888/att/api/transactionReport/"
# use General token
headers = {
    "Content-Type": "application/json",
    "Authorization": "token 9ea1dfbad18d482738bedf550c8539cff2a3a2fb",
}
data = {
    # "export_type": ".xls",
    "start_date": "2022-01-10 00:00:00",
    "end_date": "2022-01-20 07:00:00",
}


response = requests.get(url, data=json.dumps(data), headers=headers)
print(response.text)