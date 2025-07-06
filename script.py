import msal
import requests
import json
import pandas as pd
import numpy as np
import datetime
from datetime import datetime, date
import os
import time
import re

# Define the necessary variables
client_id = 'client_id'
client_secret = 'client_secret'
tenant_id = 'tenant_id'
site_name = 'site_name'
tenant_name = 'tenant_name'
sharepoint_metadata_file_path="sharepoint_metadata_excel_file_path.xlsx"
headers={}
site_id=None
list_meta=None
choices_dic=None
existing_list={}
excel_data = {}

def get_access_token():
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f'https://login.microsoftonline.com/{tenant_id}',
        client_credential=client_secret
    )
    result = app.acquire_token_for_client(scopes=['https://graph.microsoft.com/.default'])

    if 'access_token' in result:
        return result['access_token']
    else:
        print("❌ Error acquiring token:", result.get('error_description'))
        exit(-99)

def msgraph_init():
    global site_id, headers

    access_token = get_access_token()
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    }

    site_response = requests.get(
        f'https://graph.microsoft.com/v1.0/sites/{tenant_name}.sharepoint.com:/sites/{site_name}',
        headers=headers
    )

    site_data = site_response.json()

    if 'id' in site_data:
        site_id = site_data['id']
    else:
        print("Unable to get the site id")
        exit(-99)


def msgraph_get_existing_lists():
    global site_id,headers,existing_list
    # Get all lists
    lists_response = requests.get(
        f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists',
        headers=headers
    )
    lists_data = lists_response.json()
    if 'error' in lists_data:
        print('error in getting lists because {}'.format(lists_data['error']['message']))
        exit(-99)

    for lst in lists_data.get('value', []):
        existing_list[lst['displayName']]=lst['id']

def msgraph_delete_list(list_id, list_name):
    global headers, site_id
    total_deleted = 0

    items_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?$top=999'

    while items_url:
        response = requests.get(items_url, headers=headers)

        if response.status_code != 200:
            print(f"❌ Error retrieving list items: {response.text}")
            return

        items_data = response.json()
        items = items_data.get('value', [])
        if items: print(f'🔄 Deleting {len(items)} items from list: {list_name}...')

        # 🔁 Split items into batches of 20
        for i in range(0, len(items), 20):
            batch_items = items[i:i+20]
            batch_request = {"requests": []}

            for j, item in enumerate(batch_items):
                item_id = item['id']
                batch_request["requests"].append({
                    "id": str(j + 1),
                    "method": "DELETE",
                    "url": f"/sites/{site_id}/lists/{list_id}/items/{item_id}"
                })

            #Send batch delete request
            batch_response = requests.post(
                "https://graph.microsoft.com/v1.0/$batch",
                headers={**headers, "Content-Type": "application/json"},
                json=batch_request
            )

            if batch_response.status_code != 200:
                print(f"❌ Batch delete failed: {batch_response.text}")
                continue

            results = batch_response.json().get("responses", [])
            for result in results:
                if result.get("status") == 204:
                    total_deleted += 1
                else:
                    print(f"⚠️ Failed to delete item: {result}")

            # Optional: small sleep to avoid throttling
            time.sleep(0.1)

        # Pagination
        items_url = items_data.get('@odata.nextLink')

    print(f"✅ Total items deleted from list '{list_name}': {total_deleted}")

    # Delete the list
    delete_list_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}'
    delete_list_response = requests.delete(delete_list_url, headers=headers)

    if delete_list_response.status_code == 204:
        print(f"✅ Successfully deleted list {list_name} (ID: {list_id})")
    else:
        print(f"❌ Failed to delete list {list_name}: {delete_list_response.text}")

def msgraph_deletee_list(list_id,list_name):
    global headers, site_id
    total_deleted = 0

    items_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?$top=999'

    while items_url:
        response = requests.get(items_url, headers=headers)

        if response.status_code != 200:
            print(f"Error retrieving list items: {response.text}")
            return

        items_data = response.json()
        if items_data.get('value', []): print(f'Items of {list_name} will be deleted....')
        #Delete all list items
        for item in items_data.get('value', []):
            item_id = item['id']
            delete_item_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}'

            delete_item_response = requests.delete(delete_item_url, headers=headers)

            if delete_item_response.status_code == 204:
                #print(f"Successfully deleted item with ID {item_id}")
                total_deleted += 1
            else:
                print(f"Failed to delete item with ID {item_id}: {delete_item_response.text}")
        # Pagination: get the next page
        items_url = items_data.get('@odata.nextLink')
    print(f"✅ Total items deleted: {total_deleted}")
    #Delete the list after items are deleted
    delete_list_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}'
    delete_list_response = requests.delete(delete_list_url, headers=headers)

    if delete_list_response.status_code == 204:
        print(f"Successfully deleted list {list_name} with ID {list_id}")
    else:
        print(f"Failed to delete list {list_name} with ID {list_id}: {delete_list_response.text}")

def msgraph_create_list(list_def):
    global headers,site_id,existing_list,list_meta

    if list_def['displayName'] in existing_list:
        msgraph_delete_list(existing_list[list_def['displayName']],list_def['displayName'])

    create_list_response = requests.post(
        f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists',
        headers=headers,
        data=json.dumps(list_def)
    )
    create_list_data = create_list_response.json()
    if 'error' in create_list_data:
        print('error in creating {} because {}'.format(list_def['displayName'],create_list_data['error']['message']))
        exit(-99)

    list_id=create_list_data['id']
    list_meta[list_def['displayName']]['list_id']=list_id

    update_title = {
        "enforceUniqueValues": False,
        "required": False,
        "hidden": True,
        "indexed": False}

    update_title_response = requests.patch(
        f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns/Title',
        headers=headers,
        data=json.dumps(update_title)
    )
    update_column_data = update_title_response.json()
    if 'error' in update_column_data:
        print('error in updating title in {} because'.format(list_def['displayName'],create_list_data['error']['message']))
        exit(-99)

def read_lists_metadata():
    global list_meta,choices_dic
    created_lists=[]
    lists_index_df = pd.read_excel(sharepoint_metadata_file_path, sheet_name='lists_index')
    for index, row in lists_index_df.iterrows():
        if row['list_status']=='created':
            created_lists.append(row['list_name'])

    lists_choices_df = pd.read_excel(sharepoint_metadata_file_path, sheet_name='lists_choices')
    choices_dic={}
    for index, row in lists_choices_df.iterrows():
        if row['choice_id'] not in choices_dic:
            choices_dic[row['choice_id']]=[]
        choices_dic[row['choice_id']].append(row['choice_value'])

    lists_specs_df = pd.read_excel(sharepoint_metadata_file_path, sheet_name='lists_specs')
    lists_specs_df = lists_specs_df.fillna(np.nan).replace([np.nan], [None])
    list_meta={}
    lookup_lists=[]
    for index, row in lists_specs_df.iterrows():
        if row['list_name'] in created_lists:
            continue
        if row['list_name'] not in list_meta:
            list_meta[row['list_name']]={'list_id':-1,'keys_count':0,'is_lookup':'N','fields':[]}
        list_meta[row['list_name']]['fields'].append({
            'field_name':row['field_name'].strip(),
            'is_key?': row['is_key?'].strip(),
            'is_required': row['is_required?'].strip(),
            'data_type': row['data_type'].strip(),
            'size': row['size'],
            'choice_id': row['choice_id'],
            'lookup': row['lookup'],
            'allow_multiple_values_lookup': row['allow_multiple_values_lookup']
        })
        if row['is_key?']=='Y':
            list_meta[row['list_name']]['keys_count']+=1
        if   row['lookup'] is not None :
            lkp_list_name=row['lookup'].split('.')[0].strip()
            if lkp_list_name not in lookup_lists:
                lookup_lists.append(lkp_list_name)

    for lst in list_meta:
        if lst in lookup_lists:
            list_meta[lst]['is_lookup']='Y'

    for choices_id in choices_dic:
        for i in range(len(choices_dic[choices_id])):
            choices_dic[choices_id][i]=str(choices_dic[choices_id][i])

def construct_list_definition(list_name):
    global list_meta,choices_dic
    list_def={
        "displayName": list_name,
        "columns": []
    }
    for field in list_meta[list_name]['fields']:
        column_def={"name":field['field_name']}
        if field['is_key?']=='Y' and list_meta[list_name]['keys_count']==1:
            column_def['enforceUniqueValues']=True
            column_def['indexed'] = True
        if field['is_required']=='Y':
            column_def['required'] = True
        if field['lookup'] is not None:
            if field['allow_multiple_values_lookup']=='Y':
                multiple_values=True
            else:
                multiple_values=False
            lkp_column=field['lookup'].split('.')[1].strip()
            lkp_list=field['lookup'].split('.')[0].strip()
            if lkp_list in list_meta:
                lkp_list_id=list_meta[lkp_list]['list_id']
            else:
                lkp_list_id=existing_list[lkp_list]

            column_def['lookup']={"columnName": lkp_column,"listId": lkp_list_id,"allowMultipleValues": multiple_values }
        elif field['choice_id'] is not None:
            column_def["choice"]= {"choices": choices_dic[field['choice_id']]}
        else:
            if field['data_type']=='int':
                column_def['number']={"decimalPlaces":"none"}
            elif field['data_type'] in ['float','double']:
                column_def['number']={"decimalPlaces":"five"}
            elif field['data_type'] == 'string':
                maxLength=250
                allowMultipleLines=False
                linesForEditing=1
                if field['size']=='MAX':
                    maxLength=12000
                    allowMultipleLines=True
                    linesForEditing=5

                column_def['text']={'allowMultipleLines':allowMultipleLines,'maxLength':maxLength,'linesForEditing':linesForEditing
                                    ,"appendChangesToExistingText": False,"textType": "plain"}
            elif field['data_type'] in ['date','datetime']:
                format='dateOnly'
                if field['data_type']=='datetime':
                    format='dateTime'
                column_def['dateTime']={'format':format,"displayAs": "standard"}
        list_def['columns'].append(column_def)
    return list_def

def create_lists():
    global list_meta
    print(list_meta)
    #create_lookup_lists
    for list in list_meta:
        if list_meta[list]['is_lookup']=='Y':
            list_def=construct_list_definition(list)
            msgraph_create_list(list_def)

    for list in list_meta:
        if list_meta[list]['is_lookup']=='N':

            list_def=construct_list_definition(list)
            msgraph_create_list(list_def)

def read_excel_data():
    global excel_data, sharepoint_metadata_file_path, list_meta

    excel_data = {}
    try:
        xls = pd.ExcelFile(sharepoint_metadata_file_path)

        for list_name in list_meta:
            try:
                # To handle NA of Nesma Airlines code or NA of Namibia code
                if list_name== 'asq_airline_lkp' or list_name=='asq_country_lkp' or list_name=='unified_country_lkp': df = pd.read_excel(xls, sheet_name=list_name, dtype=str, keep_default_na=False)
                else :
                    df = pd.read_excel(xls, sheet_name=list_name)
                    df = df.fillna(np.nan).replace([np.nan], [None])
                records = df.to_dict(orient="records")
                excel_data[list_name] = records
            except ValueError:
                print(f"⚠️ Sheet for list '{list_name}' not found in Excel. Skipping...")
    except Exception as e:
        print(f"❌ Failed to open Excel file: {e}")
        exit(-99)

def normalize(value):
    if value is None:
        return ""
    # First, normalize invisible characters and multiple spaces
    cleaned = re.sub(r'[\u200B\u200C\u200D\uFEFF\u00A0]', '', str(value))
    cleaned = re.sub(r'[ \t]+', ' ', cleaned)  # replace multiple spaces/tabs with single space
    return cleaned.strip().lower()

def fetch_lookup_data(lookup_list_name, lookup_field):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lookup_list_name}/items?$expand=fields"
    lookup_data = {}

    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            items = data.get('value', [])
            for item in items:
                fields = item.get('fields', {})
                lookup_value = fields.get(lookup_field)
                if lookup_list_name== 'hon_device_lkp':
                    lookup_value = normalize(lookup_value)
                item_id = item.get('id')

                if lookup_value:
                    lookup_data[lookup_value] = item_id
                else:
                    print(f"[WARNING] Lookup field '{lookup_field}' not found in item with ID: {item_id}")

            # Pagination: move to next page if exists
            url = data.get('@odata.nextLink', None)
        else:
            print(f"[ERROR] Fetching lookup data for {lookup_list_name}: {response.status_code} {response.text}")
            break

    return lookup_data
'''
def get_internal_field_names(list_id):
    response = requests.get(
        f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns',
        headers=headers
    )

    columns = response.json().get("value", [])
    name_map = {}

    for column in columns:
        display_name = column.get("displayName")
        internal_name = column.get("name")
        name_map[display_name] = internal_name

    return name_map
'''
def insert_lists_init_data(list_name, records):
    global list_meta, headers, site_id, excel_data
    list_id = list_meta[list_name]['list_id']
    #internal_field_map = get_internal_field_names(list_id)

    for row in records:
            fields = {}

            for field in list_meta[list_name]['fields']:
                field_name = field['field_name']
                value = row.get(field_name)
                #internal_name = internal_field_map.get(field_name, field_name)
                #print("internal_name :",internal_name)

                #print("display_name: ",field_name)
                data_type = field.get('data_type')
                is_required = field.get('is_required') == 'Y'
                lookup = field.get('lookup')
                allow_multiple=field.get('allow_multiple_values_lookup') == 'Y'
                # --- Handle Lookup Field ---
                if lookup and value:
                    if list_name=='hon_feedback':
                        value = normalize(value)
                    lookup_list, lookup_field = lookup.split('.')
                    lookup_data = fetch_lookup_data(lookup_list, lookup_field)
                    # If it's a multi-value lookup, 'value' might be a delimiter-separated string of multiple lookup values, e.g. "Value1,Value2"
                    if allow_multiple:
                        if ',' in str(value):
                          values_list = [str(int(v.strip())) for v in str(value).split(',')]
                        else :
                            values_list =[str(int(value))]
                        lookup_ids = []
                        for val in values_list:
                            if int(val) in lookup_data:
                                lookup_ids.append(int(lookup_data[int(val)]))
                            else:
                                print(type(val))
                                print(type(lookup_data['8']))
                                print(f"⚠️ Lookup value '{val}' not found in '{lookup_list}'")
                                exit(-99)

                        # Set field to expected multi-value format
                        fields[f"{field_name}LookupId@odata.type"] = 'Collection(Edm.Int32)'
                        fields[f"{field_name}LookupId"] =  lookup_ids

                    else:

                        #print(lookup_data)
                        if value in lookup_data :
                            fields[f"{field_name}LookupId"] = int(lookup_data[value])
                        else:
                            print(f"⚠️ Lookup value '{value}' not found in '{lookup_list}'")
                            exit(-99)
                    continue

                # --- Handle Fields ---
                if pd.isna(value):
                    if is_required :
                        print(f"⚠️ {field_name} field is required")
                        exit(-99)
                    else :

                        value = None

                if data_type == 'date':
                    if isinstance(value, pd.Timestamp):
                        value = value.isoformat()
                    elif value is None:
                        value = None
                    else:
                        try:
                            value = pd.to_datetime(value).isoformat()
                        except Exception:
                            print(f"❌ Invalid date format for {field_name}: {value}")
                            value = None

                elif data_type == 'string':
                    if value is not None:
                        value = str(value)
                        '''
                        if field['size'] == 'MAX': field['size'] = 350;
                        if field['size'] and len(value) > field['size']:
                            print(f"⚠️ Invalid size ({len(value)} chars) for {field_name} that requires max:{field['size']} chars")
                            exit(-99)
                        '''
                elif data_type == 'int':
                    try:
                        value = int(value)
                    except:
                        value = None

                elif data_type  in ['float', 'double']:
                    try:
                        value = float(value)
                    except:
                        value = None

                # Convert choice values to string
                if field.get('choice_id') is not None and value is not None:
                    value = str(value)

                fields[field_name] = value

            item = {"fields": fields}

            print(f"\n📦 Final prepared item: {json.dumps(item, ensure_ascii=False)}")

            response = requests.post(
                f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items',
                headers=headers,
                data=json.dumps(item, ensure_ascii=False)
            )

            if response.status_code == 401:
                print("🔄 Token expired. Refreshing token...")
                new_token = get_access_token()
                headers['Authorization'] = f'Bearer {new_token}'

                # Retry the request
                response = requests.post(
                    f'https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items',
                    headers=headers,
                    data=json.dumps(item, ensure_ascii=False)
                )
            print(f"🔁 RESPONSE: {response.status_code} {response.text}")
            try:
                resp_json = response.json()
            except Exception:
                resp_json = {"error": {"message": response.text}}


            if response.status_code >= 400:
                error_msg = resp_json.get("error", {}).get("message", "Unknown error")
                print(f"❌ Failed to insert into '{list_name}': {error_msg}")
                exit(-99)
            else:
                item_id = resp_json.get('id')
                print(f"✅ Inserted into '{list_name}' | Item ID: {item_id}")

msgraph_init()
msgraph_get_existing_lists()
read_lists_metadata()
create_lists()
print (f"Lists meta :{list_meta}")
read_excel_data()
# Insert lookup tables first
for list_name, meta in list_meta.items():
    if meta['is_lookup'] == 'Y':
        records = excel_data.get(list_name, [])
        insert_lists_init_data(list_name, records)

# Insert non-lookup tables
for list_name, meta in list_meta.items():
    if meta['is_lookup'] != 'Y':
        records = excel_data.get(list_name, [])
        insert_lists_init_data(list_name, records)
