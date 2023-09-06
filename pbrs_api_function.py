import json
import requests
import os
from requests_ntlm2 import HttpNtlmAuth
import pandas as pd
from username_password import username, password, localhost

PowerBIReports = 'PowerBIReports'
Folders = "Folders"
Schedules = "Schedules"

baseurl = "{}/Reports/api/v2.0".format(localhost)


# Auth
def get_auth(username, password):
    auth = HttpNtlmAuth(username, password)
    return auth


auth = get_auth(username, password)


# Get
def get_df_from_pbrs(tag):
    result = requests.get(os.path.join(baseurl, tag), auth=auth).json()
    df = pd.DataFrame(result['value'])
    return df


def get_json_from_pbrs(tag):
    result = requests.get(os.path.join(baseurl, tag), auth=auth).json()
    print(result)


## Refresh Plans
def get_report_CacheRefreshPlans(id):
    baseurl_report = "{}/PowerBIReports({})".format(baseurl, id)
    result = requests.get(os.path.join(baseurl_report, 'CacheRefreshPlans'), auth=auth).json()
    df = pd.DataFrame(result['value'])
    return df


def get_CacheRefreshPlans_for_all_reports():
    re = []
    for i in range(len(get_df_from_pbrs(PowerBIReports))):
        id = get_df_from_pbrs(PowerBIReports)['Id'][i]
        df_row = get_report_CacheRefreshPlans(id)
        re.append(df_row)
    df = pd.concat(re, axis=0)
    return df


# Users
def get_report_Policies(id):
    baseurl_report = "{}/PowerBIReports({})".format(baseurl, id)
    result = requests.get(os.path.join(baseurl_report, 'Policies'), auth=auth).json()
    df = pd.DataFrame(result['Policies'])
    df['ReportID'] = id
    df = df[['ReportID', 'GroupUserName', 'Roles']]
    return df


def get_Policies_for_all_reports():
    re = []
    for i in range(len(get_df_from_pbrs(PowerBIReports))):
        id = get_df_from_pbrs(PowerBIReports)['Id'][i]
        df_row = get_report_Policies(id)
        re.append(df_row)
    df = pd.concat(re, axis=0)
    return df


# Data Role Model
def get_report_DataModelRoles(id):
    url = "{}/PowerBIReports({})".format(baseurl, id)
    result = requests.get(os.path.join(url, 'DataModelRoles'), auth=auth).json()
    df = pd.DataFrame(result['value'])
    df['ReportID'] = id
    return df


def get_report_DataModelRoleAssignments(id):
    url = "{}/PowerBIReports({})".format(baseurl, id)
    result = requests.get(os.path.join(url, 'DataModelRoleAssignments'), auth=auth).json()
    df = pd.DataFrame(result['value'])
    df['ReportID'] = id
    if len(df) != 0:
        df['DataModelRoles'] = df['DataModelRoles'].apply(lambda x: x[0])
    return df


def get_DataModelRoles_for_all_reports():
    re = []
    for i in range(len(get_df_from_pbrs(PowerBIReports))):
        id = get_df_from_pbrs(PowerBIReports)['Id'][i]
        df_row = get_report_DataModelRoles(id)
        re.append(df_row)
    df = pd.concat(re, axis=0)
    return df


def get_DataModelRoleAssignments_for_all_reports():
    re = []
    for i in range(len(get_df_from_pbrs(PowerBIReports))):
        id = get_df_from_pbrs(PowerBIReports)['Id'][i]
        df_row = get_report_DataModelRoleAssignments(id)
        re.append(df_row)
    df = pd.concat(re, axis=0)
    return df


# Post
# Rebuild all Folders
def post_get_folder_info(folders, id):
    df = folders[folders.Id == id]
    return df.to_dict(orient='records')[0]


def rebuild_all_folders(folders):
    for i in range(len(folders)):
        id = folders.Id.values[i]
        folder_data = post_get_folder_info(folders, id)
        response = requests.post(os.path.join(baseurl, "Folders"), json=folder_data, auth=auth)
        if response.status_code in [201, 204]:
            print("{} created successfully.".format(folders.Name.values[i]))
        else:
            print("Failed to create new folder. Status code:", response.status_code)


# Delete
# Delete a Folder
def delete_a_folder(id):
    response = requests.delete("{}/Folders({})".format(baseurl, id),
                               auth=auth)
    if response.status_code in [201, 204]:
        print("folder deleted successfully.")
    else:
        print("Failed to delete folder. Status code:", response.status_code)
