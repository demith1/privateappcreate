import pandas as pd
import requests
import time
import sys
import re

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(sys.argv[1])


api_token = input("API Token: ")
tenant_url = input("Complete Tenant URL(ex: tenantname.eu.goskope.com): ")
headers = { 'Netskope-Api-Token': str(api_token)}
create_app_request = 'https://'+str(tenant_url)+'/api/v2/steering/apps/private'
get_tags = 'https://'+str(tenant_url)+'/api/v2/steering/apps/private/tags'


def retrieve_tags():
    
    tag_list = []
    response = requests.get(get_tags, headers=headers)
    data = response.json()
    
    for tag in data['data']['tags']:
        tag_list.append(tag['tag_name'])

    return tag_list


# Iterate over each row in the DataFrame and make a POST request for each app

for index, row in df.iterrows():

    # Build the JSON payload for the POST request
    payload = {
        "app_name": row['App Name'],
        "host": row['Hosts'],
        "protocols": [],
        "tags": [],
        "use_publisher_dns": False,
        "clientless_access": False,
        "trust_self_signed_certs": True
    }    

    # Verify if the TCP and UDP ports exist and if they contain invalid characters
    if row['TCP Ports'] is not None and not pd.isna(row['TCP Ports']):
        if not re.match(r'^\d+(\s*,\s*\d+)*$', str(row['TCP Ports'])):
            print(f"Invalid TCP port values '{row['TCP Ports']}' for app '{row['App Name']}'. Skipping...")
            continue
        else:
            payload['protocols'].append({
                "type": "tcp",
                "port": str(row['TCP Ports'])
        })
    
    if row['UDP Ports'] is not None and not pd.isna(row['UDP Ports']):
        if not re.match(r'^\d+(\s*,\s*\d+)*$', str(row['UDP Ports'])):
            print(f"Invalid UDP port values '{row['UDP Ports']}' for app '{row['App Name']}'. Skipping...")
            continue
        else:
            payload['protocols'].append({
                "type": "udp",
                "port": str(row['UDP Ports'])
        })

    # Verify if the tag exists in the xls file and if yes, add it to the payload
    if row['Tags'] is not None and not pd.isna(row['Tags']):
        # If the tag in row['Tags'] doesnt exist in the tenant, it will be automatically created
        # If there is more than one tag, comma separated, it will create a single string tag which we don't want
        # Splitting the row based on "," to be able to create and assign individual tags
        tags = row['Tags'].split(',')
        for tag in tags:
            payload['tags'].append({
                "tag_name": str(tag)
                })

    # Payload is built, creating the private app in the tenant        
    response = requests.post(create_app_request,headers=headers,json=payload)
    
    # Check the status code of the response
    if response.status_code == 200:
        print(f"App '{row['App Name']}' created successfully.")
    else:
        print(f"Error creating app '{row['App Name']}' ({response.status_code}): {response.text}")
    time.sleep(1)
