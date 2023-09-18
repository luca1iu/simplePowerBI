from pbrs_api_function import *
import ast

# get the latest PowerBiReportServer_Info.xlsx
latest_file = get_latest_file(pbrs_excel_info_path)
# get the current working directory
cwd = os.getcwd()

if latest_file is not None:
    print('Latest file:', latest_file)
    latest_file_absolut = os.path.join(cwd, latest_file)
    try:
        # get current server info
        folders = get_df_from_pbrs('Folders')
        powerbireports = get_df_from_pbrs('PowerBIReports')
        schedules = get_df_from_pbrs('Schedules')
        refresh_plans = get_CacheRefreshPlans_for_all_reports()
        policies_user = get_Policies_for_all_reports()
        DataModelRoles = get_DataModelRoles_for_all_reports()
        DataModelRoleAssignments = get_DataModelRoleAssignments_for_all_reports()
        print("successfully: get current server info")

        # read the latest excel file
        folders_excel = pd.read_excel(latest_file_absolut, sheet_name="Folders", engine='openpyxl')
        powerbireports_excel = pd.read_excel(latest_file_absolut, sheet_name="PowerBIReports", engine='openpyxl')
        schedules_excel = pd.read_excel(latest_file_absolut, sheet_name="Schedules", engine='openpyxl')
        refresh_plans_excel = pd.read_excel(latest_file_absolut, sheet_name="CacheRefreshPlans", engine='openpyxl')
        policies_user_excel = pd.read_excel(latest_file_absolut, sheet_name="Policies", engine='openpyxl')
        DataModelRoles_excel = pd.read_excel(latest_file_absolut, sheet_name="DataModelRoles", engine='openpyxl')
        DataModelRoleAssignments_excel = pd.read_excel(latest_file_absolut,
                                                       sheet_name="DataModelRoleAssignments", engine='openpyxl')
        print("successfully: get Latest file excel info")

        # test if the current info equal to the excel info
        folders_test = folders.fillna('').drop('Roles', axis=1).equals(
            folders_excel.fillna('').drop('Roles', axis=1))
        powerbireports_test = powerbireports.fillna('').drop('Roles', axis=1).equals(
            powerbireports_excel.fillna('').drop('Roles', axis=1))
        schedules_test = schedules.fillna('').drop('Definition', axis=1).drop('NextRunTime', axis=1).drop('LastRunTime',
                                                                                                          axis=1).equals(
            schedules_excel.fillna('').drop('Definition', axis=1).drop('NextRunTime', axis=1).drop('LastRunTime',
                                                                                                   axis=1))
        refresh_plans_test = refresh_plans_excel[['Id', 'ModifiedDate']].equals(
            refresh_plans[['Id', 'ModifiedDate']].reset_index(drop=True))

        policies_user_excel.Roles = policies_user_excel.Roles.apply(lambda x: ast.literal_eval(x))
        policies_user_test = policies_user.reset_index(drop=True).equals(policies_user_excel.reset_index(drop=True))
        DataModelRoles_test = DataModelRoles_excel.equals(
            DataModelRoles) and DataModelRoleAssignments_excel.equals(
            DataModelRoleAssignments)

        # print("folders_test", folders_test)
        # print('powerbireports_test', powerbireports_test)
        # print('schedules_test', schedules_test)
        # print('refresh_plans_test', refresh_plans_test)
        # print('policies_user_test', policies_user_test)
        # print('DataModelRoles_test', DataModelRoles_test)

        # if there is no updates, do not storage any excel file
        if folders_test and powerbireports_test and schedules_test and refresh_plans_test and policies_user_test and DataModelRoles_test:
            print("no update")
        else:
            with pd.ExcelWriter(pbrs_excel, engine='openpyxl') as writer:
                # write each DataFrame to the writer object
                folders.to_excel(writer, sheet_name='Folders', index=False)
                powerbireports.to_excel(writer, sheet_name='PowerBIReports', index=False)
                schedules.to_excel(writer, sheet_name='Schedules', index=False)
                refresh_plans.to_excel(writer, sheet_name='CacheRefreshPlans', index=False)
                policies_user.to_excel(writer, sheet_name='Policies', index=False)
                DataModelRoles.to_excel(writer, sheet_name='DataModelRoles', index=False)
                DataModelRoleAssignments.to_excel(writer, sheet_name='DataModelRoleAssignments', index=False)
                print("found update and excel saved successfully")
    except Exception as e:
        print(e)
else:
    folders = get_df_from_pbrs('Folders')
    powerbireports = get_df_from_pbrs('PowerBIReports')
    schedules = get_df_from_pbrs('Schedules')
    refresh_plans = get_CacheRefreshPlans_for_all_reports()
    policies_user = get_Policies_for_all_reports()
    DataModelRoles = get_DataModelRoles_for_all_reports()
    DataModelRoleAssignments = get_DataModelRoleAssignments_for_all_reports()

    writer = pd.ExcelWriter(pbrs_excel, engine='openpyxl')

    # write each DataFrame to the writer object
    folders.to_excel(writer, sheet_name='Folders', index=False)
    powerbireports.to_excel(writer, sheet_name='PowerBIReports', index=False)
    schedules.to_excel(writer, sheet_name='Schedules', index=False)
    refresh_plans.to_excel(writer, sheet_name='CacheRefreshPlans', index=False)
    policies_user.to_excel(writer, sheet_name='Policies', index=False)
    DataModelRoles.to_excel(writer, sheet_name='DataModelRoles', index=False)
    DataModelRoleAssignments.to_excel(writer, sheet_name='DataModelRoleAssignments', index=False)

    # save the Excel file
    writer.close()
