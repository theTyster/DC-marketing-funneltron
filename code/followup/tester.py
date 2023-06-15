#! python3
# DEPRECATION WARNING: HubSpot API (HAPI) keys are being deprecated, and will be removed from use on 30 November 2022.

import hubspot
from pprint import pprint
from hubspot.crm.contacts import PublicObjectSearchRequest, ApiException

client = hubspot.Client.create(api_key="661a652d-8ac7-4471-859f-dd3fa4364dc9")

public_object_search_request = PublicObjectSearchRequest(filter_groups=[{"filters":[{"value":"***REMOVED***","propertyName":"email","operator":"EQ"}]}])

try:
    api_response = client.crm.contacts.search_api.do_search(public_object_search_request=public_object_search_request)
    pprint(api_response)

except ApiException as e:
    print("Exception when calling search_api->do_search: %s\n" % e)
