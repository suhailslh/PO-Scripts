import json
import getpass
import keyring
import requests
from bs4 import BeautifulSoup
import re
import pandas as pd

properties = dict()
with open("properties.json", "r") as f:
    properties = json.load(f)
    print(properties)

service_id = 'CC'
username = getpass.getuser()
keyring.set_password(service_id, username, getpass.getpass("Enter Password for %s: " % username))
df = pd.DataFrame()
channel_id_list = []
user_list = []

for system, host in zip(properties['system/s'].split(","), properties['host/s'].split(",")):
    system = system.strip()
    host = host.strip()
    url = properties['protocol'] + "://" + host + properties['endpoint']

    channel_list = []
    with open("ChannelQueryRequest.xml", "r") as f:
        response = requests.post(
            url, 
            data=f.read(), 
            headers={
                "SOAPAction": "query",
                "Content-Type": "text/xml"
            }, 
            auth=(username, keyring.get_password(service_id, username))
        )
        bs_data = BeautifulSoup(response.content, features="xml")  
        channel_list = list(zip(
            bs_data.find_all('PartyID'),
            bs_data.find_all('ComponentID'),
            bs_data.find_all('ChannelID')        
        ))
        print("Channel Count for " + system + ": ", len(channel_list))

    with open("ChannelReadRequest.xml", "r") as f:
        request_body = """
        <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:urn="urn:CommunicationChannelServiceVi" xmlns:urn1="urn:com.sap.aii.ibdir.server.api.types">
        <soapenv:Header/>
        <soapenv:Body>
        <urn:read>
        <urn:CommunicationChannelReadRequest>
        """
        channel_read_request_xml = f.read()
        for channel in channel_list:
            request_body = request_body + channel_read_request_xml\
            .replace("{{PartyID}}", channel[0].string or "")\
            .replace("{{ComponentID}}", channel[1].string)\
            .replace("{{ChannelID}}", channel[2].string)
        request_body += """
        </urn:CommunicationChannelReadRequest>
        </urn:read>
        </soapenv:Body>
        </soapenv:Envelope>
        """
        response = requests.post(
            url, 
            data=request_body, 
            headers={
                "SOAPAction": "query",
                "Content-Type": "text/xml"
            }, 
            auth=(username, keyring.get_password(service_id, username))
        )
        bs_data = BeautifulSoup(response.content, features="xml")  
        pattern = re.compile("httpDestination|logonUser")
        for channel in bs_data.find_all('CommunicationChannel'):
            for x in channel.find_all('AdapterSpecificAttribute'):
                if pattern.match(x.find("Name").string):
                    channel_id_list += [channel.find("ChannelID").string]
                    user_list += [str(x.find("Value").string) + "-" + system]

df["ChannelID"] = channel_id_list
df["User"] = user_list
with pd.ExcelWriter(properties['dest'], engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Sheet1", startrow=1, header=False, index=False)
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]
    (max_row, max_col) = df.shape
    column_settings = [{"header" : column} for column in df.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {"columns" : column_settings})
    worksheet.set_column(0, max_col - 1, 50)