#! python3
# DEPRECATION WARNING: HubSpot API (HAPI) keys are being deprecated, and will be removed from use on 30 November 2022.
import hubspot
from pprint import pprint
from hubspot.crm.contacts import SimplePublicObjectInput, ApiException

client = hubspot.Client.create(api_key="661a652d-8ac7-4471-859f-dd3fa4364dc9")

properties = {
    "lifecyclestage": "marketingqualifiedlead",
    "n2023_account_status": "Customer",
    "hubspot_owner_id": "97850772"
}
simple_public_object_input = SimplePublicObjectInput(properties=properties)
try:
    api_response = client.crm.contacts.basic_api.update(contact_id="6823351", simple_public_object_input=simple_public_object_input)
    pprint(api_response)
except ApiException as e:
    print("Exception when calling basic_api->update: %s\n" % e)
