import os
import requests
import pandas as pd
from dotenv import load_dotenv

load_dotenv()
ZABBIX_API_URL = os.getenv('ZABBIX_URL')
ZABBIX_TOKEN = os.getenv('ZABBIX_TOKEN')

def zabbix_api_call(method, params):
    headers = {'Content-Type': 'application/json-rpc'}
    payload = {
        "jsonrpc": "2.0",
        "method": method,
        "params": params,
        "auth": ZABBIX_TOKEN,
        "id": 1
    }
    response = requests.post(ZABBIX_API_URL, json=payload, headers=headers)
    return response.json()

def get_all_hosts():
    params = {
        "output": "extend",
        "selectTags": "extend"
    }
    result = zabbix_api_call("host.get", params)
    return result['result']

def generate_excel_file(hosts):
    df = pd.DataFrame(hosts)

    output_file = "zabbix_hosts.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Excel file '{output_file}' generated successfully.")

if __name__ == "__main__":
    hosts = get_all_hosts()
    generate_excel_file(hosts)