# This custom code is used with Zapier to automatically link a related 'Industry' field in the Insightly CRM

from pprint import pprint

input_data = {'c_id': '326342038',
              'org_id': '155115603'}

import requests as r

# Insightly Credentials
username = 'hidden'
password = ''

# Input Data
c_id = int(input_data['c_id'])
org_id = int(input_data['org_id'])


# Insightly Functions
def insightly_get(in_object, object_id):
    get_object = r.get(f'https://api.insightly.com/v3.1/{in_object}/{object_id}', auth=(username, password))
    object_json = get_object.json()
    return object_json


def insightly_put(in_object, json):
    put_object = r.put(f'https://api.insightly.com/v3.1/{in_object}', json=json, auth=(username, password))
    return put_object.status_code


def dict_search(list_name, field_name):
    return next((item['FIELD_VALUE'] for item in list_name if item['FIELD_NAME'] == field_name), None)


def dict_index(list_name, field_name):
    return next((i for i, item in enumerate(list_name) if item['FIELD_NAME'] == field_name), None)


# Define fields in Contact
contact = insightly_get('Contacts', c_id)
c_cf = contact['CUSTOMFIELDS']
c_owner = contact['OWNER_USER_ID']
c_industry = dict_search(c_cf, 'Industry2__c')
c_user_responsible = dict_search(c_cf, 'User_Responsible__c')

# Define fields in Organization
org = insightly_get('Organisation', org_id)
org_cf = org['CUSTOMFIELDS']
org_owner = org['OWNER_USER_ID']
o_industry = dict_search(org_cf, 'Industry__c')
org_user_responsible = dict_search(org_cf, 'User_Responsible__c')


# If all the information is the same, don't update!
if c_owner == org_owner and c_industry == o_industry and c_user_responsible == org_user_responsible:
    print('Owner, industry, and user responsible are the same!')
    # owner_industry_user_repsonsible_are_the_same

# Industry field
if c_industry is not None:
    industry_index = dict_index(c_cf, 'Industry2__c')
    c_cf[industry_index]['FIELD_VALUE'] = o_industry

else:
    new_industry_cf = {'CUSTOM_FIELD_ID': 'Industry2__c',
            'FIELD_NAME': 'Industry2__c',
            'FIELD_VALUE': o_industry}

    c_cf.append(new_industry_cf)

# User Responsible field
if c_user_responsible is not None:
    user_responsible_index = dict_index(c_cf, 'User_Responsible__c')
    c_cf[user_responsible_index]['FIELD_VALUE'] = org_user_responsible
else:
    new_user_responsible_cf = {'CUSTOM_FIELD_ID': 'User_Responsible__c',
            'FIELD_NAME': 'User_Responsible__c',
            'FIELD_VALUE': org_user_responsible}

    c_cf.append(new_user_responsible_cf)

json = {'CONTACT_ID': c_id,
        'OWNER_USER_ID': org_owner,
        'CUSTOMFIELDS': c_cf}

update_c = insightly_put('Contacts', json=json)
