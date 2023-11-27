from hubspot import HubSpot
import sys
# python3 <filename>  <token> 
# https://github.com/HubSpot/hubspot-api-python   
n = len(sys.argv)
if n != 2:
    print('Please provide valid token')
    exit(0)
token = sys.argv[1]
api_client = HubSpot(access_token=token)
# token  (pat-na1-374fdf1d-2b93-43e8-b0c5-5d83f42f1e9b)
# or set your access token later
api_client = HubSpot()
api_client.access_token = token
all_contacts = api_client.crm.contacts.get_all()
all_deals = api_client.crm.deals.get_all()
all_companies = api_client.crm.companies.get_all()
print(all_companies)
