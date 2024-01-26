""" Copyright (c) 2022 Cisco and/or its affiliates.
This software is licensed to you under the terms of the Cisco Sample
Code License, Version 1.1 (the "License"). You may obtain a copy of the
License at
           https://developer.cisco.com/docs/licenses
All use of the material herein must be in accordance with the terms of
the License. All rights not expressly granted by the License are
reserved. Unless required by applicable law or agreed to separately in
writing, software distributed under the License is distributed on an "AS
IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
or implied.
"""

#Imports
import os
import requests
import json
from dotenv import load_dotenv
import xlsxwriter
import pandas as pd

#Environment Variables
load_dotenv()
MERAKI_BASE_URL = os.environ['MERAKI_BASE_URL']


#Headers
headers = {
    "Content-Type": "application/json",
    "Accept": "application/json",
    "X-Cisco-Meraki-API-Key": os.environ['API_KEY']
}



#Get the IDs of all organizations
def get_organizations_names_ids():
    try:
        organizations = requests.get(MERAKI_BASE_URL+'/organizations', headers=headers).json()
        for org in organizations:
            orgId=org["id"]
            orgName=org["name"]
            get_org_devices(orgId,orgName)
    except Exception as e:
            print("Exception in get_organizations_names_ids: " + str(e))

#Get all devices in an organization
def get_org_devices(org_id,orgName):
    try:
        switches=[]
        cameras=[]
        accessPoints=[]
        url= MERAKI_BASE_URL + f"/organizations/{org_id}/devices/statuses"
        devices = requests.get(url,headers=headers).json()

        for device in devices:
            if device["productType"] == "switch":
                switches.append(device)
            elif device["productType"] =="wireless":
                accessPoints.append(device)
            elif device["productType"] =="camera":
                cameras.append(device)
        
        switches=get_switch_active_ports(switches)
        accessPoints=get_wireless_details(org_id,accessPoints)
        create_file(switches, accessPoints, cameras, orgName)
        create_final_report()

    except Exception as e:
            print("Exception in get_org_devices: " + str(e))

def get_switch_active_ports(switches):
    try:
        for switch in switches:
            url= MERAKI_BASE_URL + f"/devices/{switch['serial']}/switch/ports/statuses"
            ports = requests.get(url,headers=headers).json()
            enabledPortsId=[]
            connectedPortsIds=[]
            for port in ports:
                port["enabled"] = str(port["enabled"])
                if port["enabled"] =="True":    
                    enabledPortsId.append(port["portId"])
                    if port["status"] =="Connected":
                        connectedPortsIds.append(port["portId"])
            enabled_ports_ids=','.join([item for item in enabledPortsId])
            switch["enabledPorts"]=enabled_ports_ids
            connected_ports_ids=','.join([item for item in connectedPortsIds])
            switch["connectedPorts"]=connected_ports_ids
            # print(json.dumps(switches, indent=2))
        return switches
    except Exception as e:
            print("Exception in get_switch_active_ports: " + str(e))



def get_wireless_details(orgId, accessPoints):
    try:
        for ap in accessPoints:
            url= MERAKI_BASE_URL + f"/organizations/{orgId}/wireless/devices/channelUtilization/byDevice?serials[]={ap['serial']}"
            utilization = requests.get(url,headers=headers).json()
            # print(json.dumps(utilization,indent=2))
            for u in utilization:
                if u["byBand"]:
                    for band in u["byBand"]:
                        if band["band"]=="2.4":
                            ap["2.4 wifi"]=band["wifi"]["percentage"]
                            ap["2.4 nonWifi"]=band["nonWifi"]["percentage"]
                            ap["2.4 total"]=band["total"]["percentage"]
                        else:
                            ap["5 wifi"]=band["wifi"]["percentage"]
                            ap["5 nonWifi"]=band["nonWifi"]["percentage"]
                            ap["5 total"]=band["total"]["percentage"]
                else:
                    ap["2.4 wifi"]=""
                    ap["2.4 nonWifi"]=""
                    ap["2.4 total"]=""
                    ap["5 wifi"]=""
                    ap["5 nonWifi"]=""
                    ap["5 total"]=""
        
            url= MERAKI_BASE_URL + f"/devices/{ap['serial']}/clients"
            clients = requests.get(url,headers=headers).json()
            ap["clients"]=str(len(clients))

        return accessPoints
    except Exception as e:
            print("Exception in get_wireless_details: " + str(e))


def create_file(switches, accessPoints, cameras, orgName):
    workbook = xlsxwriter.Workbook(f'./Reports/{orgName}.xlsx')
    worksheet = workbook.add_worksheet(f"{orgName}")
    row = 0
    column = 0
    merge_format = workbook.add_format(
    {
        "bold": 1,
        "align": "center",
        "valign": "vcenter"
       
        }
    )
    headers=["Name", "Model","MAC","Serial", "LAN IP", "Status","Enabled Ports","Connected Ports"]
    worksheet.merge_range("A2:H2", "Switches",merge_format)
    row=2
    for header in headers:
        worksheet.write(row, column, header)
        column += 1
    
    column=0
    row=3
    for switch in switches:
        content =[switch["name"],switch["model"],switch["mac"],switch["serial"],switch["lanIp"],switch["status"], switch["enabledPorts"],switch["connectedPorts"]]
        for item in content :
            worksheet.write(row, column, item)
            column += 1
        row+=1
        column=0
    
    row=row+3
    headers=["Name", "Model","MAC","Serial", "LAN IP", "Status"]
    worksheet.merge_range(f"A{row}:H{row}", "Cameras",merge_format)
    
    for header in headers:
        worksheet.write(row, column, header)
        column += 1
    
    column=0
    row=row+1
    for camera in cameras:
        content =[camera["name"],camera["model"],camera["mac"],camera["serial"],camera["lanIp"],camera["status"]]
        for item in content :
            worksheet.write(row, column, item)
            column += 1
        row+=1
        column=0
    

    row=row+3
    headers=["Name", "Model","MAC","Serial", "LAN IP", "Status","2.4 wifi","2.4 nonWifi","2.4 total","5 wifi","5 nonWifi","5 total","Clients"]
    worksheet.merge_range(f"A{row}:H{row}", "Access Points",merge_format)
    
    for header in headers:
        worksheet.write(row, column, header)
        column += 1
    
    column=0
    row=row+1
    for ap in accessPoints:
        content =[ap["name"],ap["model"],ap["mac"],ap["serial"],ap["lanIp"],ap["status"],ap["2.4 wifi"],ap["2.4 nonWifi"],ap["2.4 total"],ap["5 wifi"],ap["5 nonWifi"],ap["5 total"],ap["clients"]]
        for item in content :
            worksheet.write(row, column, item)
            column += 1
        row+=1
        column=0
    
    workbook.close()


def create_final_report():
    output_excel = "Full-Report.xlsx"
    #List all excel files in folder
    excel_folder= "./Reports"
    excel_files = [os.path.join(root, file) for root, folder, files in os.walk(excel_folder) for file in files if file.endswith(".xlsx")]

    with pd.ExcelWriter(output_excel) as writer:
        for excel in excel_files: #For each excel
            sheet_name = pd.ExcelFile(excel).sheet_names[0] #Find the sheet name
            df = pd.read_excel(excel) #Create a dataframe
            df.to_excel(writer, sheet_name=sheet_name, index=False,merge_cells=True,header=False) #Write it to a sheet in the output excel
        
if __name__ == '__main__':
    get_organizations_names_ids()