import json
import getpass
import keyring
import requests
from bs4 import BeautifulSoup
import pandas as pd

properties = dict()
with open("properties.json", "r") as f:
    properties = json.load(f)
    print(properties)

service_id = 'SI'
username = getpass.getuser()
keyring.set_password(service_id, username, getpass.getpass("Enter Password for %s: " % username))
df = pd.DataFrame()
si_list = []
swc_list = []
system_line_list = []

for system, host in zip(properties['system/s'].split(","), properties['host/s'].split(",")):
    system = system.strip()
    host = host.strip()
    origin = properties['protocol'] + "://" + host
    url = origin + properties['endpoint']

    xml_request = ""
    with open("SIReadRequest.xml", "r") as f:
        xml_request = f.read()

    with open("url-encoded-form-data.json", "r") as f:
        form_data = json.load(f)
        form_data['userL'] = username.lower()
        form_data['queryRequestXMLL'] = xml_request
        form_data['request'] = xml_request
        response = requests.post(
            url, 
            data=form_data, 
            headers={
                'Host': host,
                'Origin': origin,
                'Content-Type': 'application/x-www-form-urlencoded'
            }, 
            auth=(username, keyring.get_password(service_id, username))
        )
        bs_data = BeautifulSoup(response.content, "html.parser")  
        xml_response = BeautifulSoup(bs_data.find("textarea", {"name" : "response"}).string, features="xml")
        rows = xml_response.find_all("r")
        print("Proxy Count for " + system + ": ", len(rows))
        with open("system-line-map.json") as _f:
            sys_line_map = json.load(_f)
            for row in rows:
                si_list += [row.find("key", {"typeID" : "ifmmessif"}).find("elem").string]
                swc = row.find("vc", {"vcType" : "S"})['caption']
                swc_list += [swc]
                system_line_list += [sys_line_map.get(swc, "NA") + "-" + system]

df['Technical Name'] = si_list
df['Type'] = ['ABAP Proxy'] * len(si_list)
df['SWC'] = swc_list
df['System Line'] = system_line_list
with pd.ExcelWriter(properties['dest'], engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Sheet1", startrow=1, header=False, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]
    (max_row, max_col) = df.shape
    column_settings = [{"header" : column} for column in df.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {"columns" : column_settings})
    worksheet.set_column(0, max_col - 1, 50)