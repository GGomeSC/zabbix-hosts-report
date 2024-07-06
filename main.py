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
        "selectTags": "extend",
        "selectParentTemplates": ["templateid", "name"]
    }
    result = zabbix_api_call("host.get", params)
    return result['result']

def get_template_tags(template_id):
    params = {
        "output": "extend",
        "templateids": template_id,
        "selectTags": "extend"
    }
    result = zabbix_api_call("template.get", params)
    if 'result' in result and result['result']:
        return result['result'][0].get('tags',[])
    else:
        print(f"Error fetching tags for template ID {template_id}: {result}")
    return []

def generate_excel_file(hosts):
    # Convert the list of hosts to a pandas DF
    host_data = []
    for host in hosts:
        host_info = {
            'HostID': host['hostid'],
            'Host': host['host'],
            'Name': host['name'],
            'Status': 'Enabled' if host['status'] == '0' else 'Disabled'
        }
        
        host_tags = host['tags']
        template_tags= []
        for template in host['parentTemplates']:
            template_tags.extend(get_template_tags(template['templateid']))

        all_tags = {f"{tag['tag']}:{tag['value']}" for tag in (host_tags + template_tags)}

        host_info['Tags'] = ', '.join(all_tags)
        host_data.append(host_info)

    df = pd.DataFrame(host_data)
    output_file = "zabbix_hosts.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Excel file '{output_file} gererated successfully.")

if __name__ == "__main__":
    hosts = get_all_hosts()
    generate_excel_file(hosts)