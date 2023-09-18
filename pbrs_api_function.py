import json
import numpy as np
import requests
import os
from requests_ntlm2 import HttpNtlmAuth
import pandas as pd
import datetime

# 服务器账户
username = "username"
password = 'password'

# web server
localhost = "localhost "
baseurl = "{}/Reports/api/v2.0".format(localhost)

# power bi gateway account
db_username = 'db_username'
db_password = 'db_password'

# folder path
today_date = format(datetime.datetime.today(), '%Y%m%d')
pbix_folder_path = r'pbix_files'
pbrs_excel_info_path = r'PowerBiReportSever_Info'
pbrs_excel = r'PowerBiReportSever_Info\PowerBiReportSever_Info-{}.xlsx'.format(today_date)


# Auth
def get_auth(username=username, password=password):
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
    for i in range(len(get_df_from_pbrs('PowerBIReports'))):
        id = get_df_from_pbrs('PowerBIReports')['Id'][i]
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
    for i in range(len(get_df_from_pbrs('PowerBIReports'))):
        id = get_df_from_pbrs('PowerBIReports')['Id'][i]
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
    for i in range(len(get_df_from_pbrs('PowerBIReports'))):
        id = get_df_from_pbrs('PowerBIReports')['Id'][i]
        df_row = get_report_DataModelRoles(id)
        re.append(df_row)
    df = pd.concat(re, axis=0)
    return df


def get_DataModelRoleAssignments_for_all_reports():
    re = []
    for i in range(len(get_df_from_pbrs('PowerBIReports'))):
        id = get_df_from_pbrs('PowerBIReports')['Id'][i]
        df_row = get_report_DataModelRoleAssignments(id)
        re.append(df_row)
    df = pd.concat(re, axis=0)
    return df


# Delete
# Delete a Folder
def delete_a_folder(id):
    response = requests.delete("{}/Folders({})".format(baseurl, id),
                               auth=auth)
    if response.status_code in [201, 204]:
        print("folder deleted successfully.")
    else:
        print("Failed to delete folder. Status code:", response.status_code)


def get_info_from_recovery_excel(df, id):
    df2 = df.replace(np.nan, None)[df.Id == id]
    return df2.to_dict(orient='records')[0]


def rebuild_all(df, tag):
    for i in range(len(df)):
        id = df.Id.values[i]
        id_data = get_info_from_recovery_excel(df, id)
        if tag in ['Folders']:
            id_data.pop("Roles")
            print(id_data)
        elif tag in ['CacheRefreshPlans']:
            id_data['Schedule'] = eval(id_data['Schedule'].replace("'", "\""))
            id_data.pop('ParameterValues')
        elif tag in ['PowerBIReports']:
            id_data.pop("Roles")
            id_data.pop("Content")
        response = requests.post(os.path.join(baseurl, tag), json=id_data, auth=auth)
        if response.status_code in [201, 204, 200]:
            print("{} created successfully.".format(df.Id.values[i]))
        else:
            print("Failed to rebuild {}. Status code:{}".format(df.Id.values[i], response.status_code))
            print("please rebuild {} manually".format(df.Id.values[i]))


def rebuild_single(df, id, tag):
    id_data = get_info_from_recovery_excel(df, id)
    if tag in ['Folders']:
        id_data.pop("Roles")
    elif tag in ['CacheRefreshPlans']:
        id_data['Schedule'] = eval(id_data['Schedule'].replace("'", "\""))
        id_data.pop('ParameterValues')
    res = requests.post(os.path.join(baseurl, tag), json=id_data, auth=auth)
    if res.status_code in [201, 204]:
        print("{} created successfully.".format(id))
    else:
        print("Failed to rebuild {}. Status code:{}".format(id, res.status_code))
        print("please rebuild {} manually".format(id))


# policies
def rebuild_policies_item_policies(df, name):
    item_policies = {}
    item_policies['InheritParentPolicy'] = False
    df_specific = df[df.Name == name][['GroupUserName', 'Roles']]
    df_specific['Roles'] = df_specific['Roles'].apply(lambda x: eval(x.replace("'", "\"")))
    item_policies['Policies'] = df_specific.to_dict('records')
    return item_policies


def rebuild_all_policies(df):
    set_report_name = list(set(df.Name.values))
    for report_name in set_report_name:
        report_id = df[df.Name == report_name]['Id'].values[0]
        item_policies = rebuild_policies_item_policies(df, name=report_name)
        res = requests.put(url='{}/PowerBIReports({})/Policies'.format(baseurl, report_id), json=item_policies,
                           auth=auth)
        if res.status_code in [200, 201]:
            print("{} created successfully.".format(report_name))
        else:
            print("Failed to rebuild {}. Status code:{}".format(report_name, res.status_code))
            print("please rebuild {} manually".format(report_name))


# upload power bi files
def upload_pbix(df, report_name, reports_path):
    report_id = df[df.Name == report_name].Id.values[0]
    file_path = os.path.join(reports_path, report_name + '.pbix')

    cwd = os.getcwd()
    file_path_absolute = os.path.join(cwd, file_path)

    url = '{}/PowerBIReports({})/Model.Upload'.format(baseurl, report_id)

    # Open the file and send a POST request to the API endpoint
    with open(file_path_absolute, 'rb') as f:
        response = requests.post(url=url, files={'File': f}, auth=auth)

    # Check the response status code
    if response.status_code in [200, 201]:
        print('File {} uploaded successfully!'.format(report_name))
    else:
        print(response.status_code)
        print('Error uploading file {}:'.format(report_name))


# add data source
def add_data_source_connection(report_id, db_username, db_password):
    data_source_id = \
        requests.get('{}/PowerBIReports({})/DataSources'.format(baseurl, report_id), auth=auth).json()['value'][0]['Id']
    data = [{"Name": None, "Description": None, "Hidden": False, "Path": "", "IsEnabled": True,
             "DataSourceSubType": "DataModel",
             "DataModelDataSource": {"Type": "Import", "Kind": "Odbc", "AuthType": "UsernamePassword",
                                     "SupportedAuthTypes": ["UsernamePassword", "Anonymous", "Windows"],
                                     "Username": db_username, "Secret": db_password,
                                     "ModelConnectionName": "84AD9DAA91BD4F0D309846E37950A68511832F89FFC9E16C667378A9BF9575A8"},
             "IsReference": False, "DataSourceType": "Odbc", "ConnectionString": "dsn=Oracle_ODAC",
             "IsConnectionStringOverridden": True, "CredentialRetrieval": "store",
             "CredentialsInServer": {"UserName": db_username, "Password": db_password, "UseAsWindowsCredentials": True,
                                     "ImpersonateAuthenticatedUser": False}, "CredentialsByUser": None,
             "Id": data_source_id}]
    url = 'http://shzedc-mws-pbi1:8080/Reports/api/v2.0/PowerBIReports({})/DataSources'.format(report_id)
    response = requests.patch(url=url, json=data, auth=auth)

    # Check the response status code
    if response.status_code in [200, 201]:
        print('Data Source {} uploaded successfully!'.format(report_id))
    else:
        print(response.status_code)
        print('Error Data Source {}:'.format(report_id))


# update DataModelRoleAssignments
def update_DataModelRoleAssignments(df_rls, report_id):
    url = '{}/PowerBIReports({})/DataModelRoleAssignments'.format(baseurl, report_id)
    df_specific = df_rls[df_rls['New_Report_Id'] == report_id][['GroupUserName', 'DataModelRoles']]
    data = df_specific.to_dict('records')
    response = requests.put(url=url, json=data, auth=auth)
    if response.status_code == 200:
        print("Row level security - {} 创建成功".format(report_id))
    else:
        print("Row level security 创建失败", response.status_code)
        print("data: {}".format(report_id))


# get the latest PowerBiReportServer_Info
def get_latest_file(path):
    latest_file = None
    latest_date = None

    for filename in os.listdir(path):
        full_path = os.path.join(path, filename)
        if os.path.isfile(full_path) and filename.endswith('.xlsx'):
            date_str = filename[-13:-5]
            date = datetime.datetime.strptime(date_str, '%Y%m%d')
            if latest_date is None or date > latest_date:
                latest_file = full_path
                latest_date = date

    return latest_file


# refresh dashboard
def refresh_dashboard(refresh_lan_id):
    url = "{}/CacheRefreshPlans({})/Model.Execute".format(baseurl, refresh_lan_id)
    res = requests.post(url, auth=auth)
    if res.status_code == 204:
        print("refresh successfully")
    else:
        print("refresh error:", res.status_code)
