"""
Sample code for using MS graph api python sdk
Install dependencies
python3 -m pip install azure-identity
python3 -m pip install msgraph-core

created by Aditya Gholba
"""
from msgraph.core import GraphClient
from azure.identity import ClientSecretCredential
import csv

values=[] #Global value incase of next data loop. 
def main():
    print('Python Graph Sample code\n')
    #App reg details 
    clientId = 'Client ID'
    clientSecret = 'Client Secret'
    tenantId = 'Tenant ID'
    authTenant = 'Tenant ID'
    graphUserScopes = 'Scopes'# 'User.Read Mail.Read Mail.Send'

    clientcredential = ClientSecretCredential(tenantId,clientId,clientSecret)
    client = GraphClient(credential=clientcredential)

    request_url = '/users/admin@M365x58245545.onmicrosoft.com/joinedTeams' #just enter the url from the resource ownwards
    users = get_users(client,request_url)
    print(users)
    jsonToCsv(users)

def get_users(client,request_url):
    global values 
    users_response = client.get(request_url)
    users = users_response.json() 
    if 'value' in users:
        values = values + users['value'] # values is a list of dict
    else:
        values = users    
    #retrive/print a particular field from the json response eg: userPrincipalName,displayName etc.
    #for user in values:
       #print(user['userPrincipalName'])

    #If the response has next data link (More than 1k records)   
    if '@odata.nextLink' in users:
        more_data = '@odata.nextLink' in users
        print(users['@odata.nextLink'])
        get_users(client,users['@odata.nextLink']) #Recursively add next data to values list.
        
    return values     
    
          
#export response as CSV
def jsonToCsv(json):
    loc2=('C:\\Users\\adgholba\\Downloads\\jsonToCsv.csv') #Enter the correct path
    with open(loc2, 'w', newline='') as f: 
        writer = csv.writer(f)
        writer.writerow(json[0])
        for row in json:
            writer.writerow(row.values())
        
       
main()            
