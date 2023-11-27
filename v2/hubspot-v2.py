import requests
import json
from datetime import datetime
import xlwt 
from xlwt import Workbook 
import sys
# python3 <filename>  <token>    
n = len(sys.argv)
if n != 2:
    print('Please provide valid token')
    exit(0)
token = sys.argv[1]
print("Start of pulling hubspot data")
headers = {'Authorization': 'Bearer ' + token}
baseUrl = 'https://api.hubapi.com/'
print("Start of contact data")

url = baseUrl + 'crm/v3/objects/contacts?limit=100'
wb = Workbook() 
style_string = "font: bold on; borders: bottom dashed"
style = xlwt.easyxf(style_string)
contacts = wb.add_sheet('CONTACTS') 
contacts.write(0, 0, "ID",style=style)
contacts.write(0, 1, "FIRST_NAME",style=style)
contacts.write(0, 2, "LAST_NAME",style=style)
contacts.write(0, 3, "EMAIL",style=style)
contacts.write(0, 4, "CREATED_AT",style=style)
contacts.write(0, 5, "UPDATED_AT",style=style)
contactsRow = 1  
contactRemain = True
while  contactRemain:
    response = requests.get(url,headers=headers,timeout=300)
    result =  response.json()
    
    if response.status_code == 200:
        if "paging" in result:
            after = None
            for contact in result["results"]:
                if "id" in contact:
                    contacts.write(contactsRow, 0, contact["id"])
                if "firstname" in contact["properties"]:
                    contacts.write(contactsRow, 1, contact["properties"]["firstname"])
                if "lastname" in contact["properties"]:
                    contacts.write(contactsRow, 2, contact["properties"]["lastname"])
                if "email" in contact["properties"]:
                    contacts.write(contactsRow, 3, contact["properties"]["email"])
                if "createdAt" in contact:
                    contacts.write(contactsRow, 4, contact["createdAt"])
                if "updatedAt" in contact:
                    contacts.write(contactsRow, 5, contact["updatedAt"])
                contactsRow +=1
            after =   result['paging']['next']['after']
            if after is None:
                break
            else:
                url = baseUrl + 'crm/v3/objects/contacts?limit=100&after='+after
                print(url)
        else:
            for contact in result["results"]:
                if "id" in contact:
                    contacts.write(contactsRow, 0, contact["id"])
                if "firstname" in contact["properties"]:
                    contacts.write(contactsRow, 1, contact["properties"]["firstname"])
                if "lastname" in contact["properties"]:
                    contacts.write(contactsRow, 2, contact["properties"]["lastname"])
                if "email" in contact["properties"]:
                    contacts.write(contactsRow, 3, contact["properties"]["email"])
                if "createdAt" in contact:
                    contacts.write(contactsRow, 4, contact["createdAt"])
                if "updatedAt" in contact:
                    contacts.write(contactsRow, 5, contact["updatedAt"])
                contactsRow +=1
            break

    else:
        print("Request got failed for url :",url)	
        print(result["message"])
        break 
print("End of contact data")


print("Start of deals")
deals = wb.add_sheet('DEALS') 
deals.write(0, 0, "ID",style=style)
deals.write(0, 1, "DEAL_NAME",style=style)
deals.write(0, 2, "DEAL_STAGE",style=style)
deals.write(0, 3, "PIPELINE",style=style)
deals.write(0, 4, "AMOUNT",style=style)
deals.write(0, 5, "CREATED_AT",style=style)
deals.write(0, 6, "CLOSE_DATE",style=style)
deals.write(0, 7, "UPDATED_AT",style=style)

dealsRow = 1
dealsUrl = baseUrl + 'crm/v3/objects/deal?limit=100'
dealRemain = True
while dealRemain:
    dealResponse = requests.get(dealsUrl,headers=headers,timeout=300)
    dealResult =  dealResponse.json()
    
    if dealResponse.status_code == 200:
        if "paging" in result:
            dealAfter= None
            for deal in dealResult["results"]:
                if "id" in deal:
                    deals.write(dealsRow, 0, deal["id"])
                if "dealname" in deal["properties"]:
                    deals.write(dealsRow, 1, deal["properties"]["dealname"])
                if "dealstage" in deal["properties"]:
                    deals.write(dealsRow, 2, deal["properties"]["dealstage"])
                if "pipeline" in deal["properties"]:
                    deals.write(dealsRow, 3, deal["properties"]["pipeline"])
                if "amount" in deal["properties"]:
                    deals.write(dealsRow, 4, deal["properties"]["amount"])
                if "createdate" in deal["properties"]:
                    deals.write(dealsRow, 5, deal["properties"]["createdate"])
                if "closedate" in deal["properties"]:
                    deals.write(dealsRow, 6, deal["properties"]["closedate"])
                if "updatedAt" in deal:
                    deals.write(dealsRow, 7, deal["updatedAt"])
                dealsRow +=1
            dealAfter =   dealResult['paging']['next']['after']
            if dealAfter is None:
                break
            else:
                url = baseUrl + 'crm/v3/objects/deal?limit=100&after='+dealAfter
                print(url)
        else:
            for deal in dealResult["results"]:
                if "id" in deal:
                    deals.write(dealsRow, 0, deal["id"])
                if "dealname" in deal["properties"]:
                    deals.write(dealsRow, 1, deal["properties"]["dealname"])
                if "dealstage" in deal["properties"]:
                    deals.write(dealsRow, 2, deal["properties"]["dealstage"])
                if "pipeline" in deal["properties"]:
                    deals.write(dealsRow, 3, deal["properties"]["pipeline"])
                if "amount" in deal["properties"]:
                    deals.write(dealsRow, 4, deal["properties"]["amount"])
                if "createdate" in deal["properties"]:
                    deals.write(dealsRow, 5, deal["properties"]["createdate"])
                if "closedate" in deal["properties"]:
                    deals.write(dealsRow, 6, deal["properties"]["closedate"])
                if "updatedAt" in deal:
                    deals.write(dealsRow, 7, deal["updatedAt"])
                dealsRow +=1
            break

    else:
        print("Request got failed for url :",dealsUrl)	
        print(dealResult["message"])
        break
print("End of deals")

print("Start of companies")
companies = wb.add_sheet('COMPANIES') 
companies.write(0, 0, "NAME",style=style)
companies.write(0, 1, "DOMAIN",style=style)
companies.write(0, 2, "INDUSTRY",style=style)
companies.write(0, 3, "CITY",style=style)
companies.write(0, 4, "PHONE",style=style)
companies.write(0, 5, "STATE",style=style)
companies.write(0, 6, "CREATED_DATE",style=style)

companiesRow = 1
companiesUrl = baseUrl + 'crm/v3/objects/companies'
companyRemain = True
while companyRemain:
    companiesResponse = requests.get(companiesUrl,headers=headers,timeout=300)
    companiesResult =  companiesResponse.json()
   
    if companiesResponse.status_code == 200:
        if "paging" in companiesResult:
            companyAfter= None
            for company in companiesResult["results"]:
                if "name" in company["properties"]:
                    companies.write(companiesRow, 0, company["properties"]["name"])
                if "domain" in company["properties"]:
                    companies.write(companiesRow, 1, company["properties"]["domain"])
                if "industry" in company["properties"]:
                    companies.write(companiesRow, 2, company["properties"]["industry"])
                if "city" in company["properties"]:
                    companies.write(companiesRow, 3, company["properties"]["city"])
                if "phone" in  company["properties"]:
                    companies.write(companiesRow, 4, company["properties"]["phone"])
                if "state" in  company["properties"]:
                    companies.write(companiesRow, 5, company["properties"]["state"])
                if "createdAt" in company:
                    companies.write(companiesRow, 6, company["createdAt"])
                companiesRow +=1
            companyAfter =   companiesResult['paging']['next']['after']
            if companyAfter is None:
                break
            else:
                companiesUrl = baseUrl + 'crm/v3/objects/companies?limit=100&after='+companyAfter
                print(companiesUrl)
        else:
            for company in companiesResult["results"]:
                if "name" in company["properties"]:
                    companies.write(companiesRow, 0, company["properties"]["name"])
                if "domain" in company["properties"]:
                    companies.write(companiesRow, 1, company["properties"]["domain"])
                if "industry" in company["properties"]:
                    companies.write(companiesRow, 2, company["properties"]["industry"])
                if "city" in company["properties"]:
                    companies.write(companiesRow, 3, company["properties"]["city"])
                if "phone" in  company["properties"]:
                    companies.write(companiesRow, 4, company["properties"]["phone"])
                if "state" in  company["properties"]:
                    companies.write(companiesRow, 5, company["properties"]["state"])
                if "createdAt" in company:
                    companies.write(companiesRow, 6, company["createdAt"])
                companiesRow +=1
            break
    else:
        print("Request got failed for url :",companiesUrl)	
        print(companiesResult["message"])
        break

print("End of companies")

file = datetime.now().strftime("%d-%m-%Y %H:%M:%S") + "-HubspotData.xls"
wb.save(file)
print("End of pulling hubspot data")
