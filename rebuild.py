from pbrs_api_function import *
from read_recovery_excel import *

try:
    try:
        # Step1: Rebuild Folders
        rebuild_all(df=folders, tag="Folders")
        print("Step 1: created folder successfully")
    except Exception as e:
        print("Step 1: created folder failed", e)

    try:
        # Step 2: Upload all pbix files
        rebuild_all(df=powerbireports, tag='PowerBIReports')
        for ReportName in powerbireports.Name.values:
            upload_pbix(df=get_df_from_pbrs('PowerBIReports'), report_name=ReportName,
                        reports_path=pbix_folder_path)
        print("Step 2: upload pbix file successfully")
    except Exception as e:
        print("Step 3: 上传pbix文件失败", e)

    try:
        # Step3: DataSource
        current_reports = get_df_from_pbrs('PowerBIReports')
        for ReportID in current_reports.Id.values:
            add_data_source_connection(report_id=ReportID, db_username=db_username, db_password=db_password)
        print("Step 3: 创建Data Source成功")
    except Exception as e:
        print("Step 3: 创建Data Source失败", e)

    try:
        # Step 4：Refresh Plans
        rebuild_all(df=refresh_plans, tag='CacheRefreshPlans')
        print("Step 4: 创建Refresh Plans成功")
    except Exception as e:
        print("Step 4: 创建Refresh Plans失败", e)

    try:
        # Step 5: policies_user
        df_report_users = policies_user.merge(powerbireports[['Id', 'Name']], left_on='ReportID',
                                              right_on='Id').drop(
            'Id',
            axis=1)
        df_report_users = df_report_users[['Name', 'GroupUserName', 'Roles']]
        current_report = get_df_from_pbrs('PowerBIReports')
        df_report_users = df_report_users.merge(current_report[['Name', 'Id']], on='Name')

        ## 给Report添加用户权限
        rebuild_all_policies(df=df_report_users)
        print("Step 5: 创建用户权限成功")
    except Exception as e:
        print("Step 5: 创建用户权限失败", e)

    try:
        # Step 6: DataModelRoleAssignments
        current_reports = get_df_from_pbrs('PowerBIReports')
        df_rls = DataModelRoleAssignments.merge(powerbireports[['Id', 'Name']], left_on='ReportID',
                                                right_on='Id').drop(
            'Id',
            axis=1)
        df_rls = df_rls.merge(current_reports[['Name', 'Id']]).rename(columns={'Id': 'New_Report_Id'})
        df_rls['DataModelRoles'] = df_rls['DataModelRoles'].apply(lambda x: [x])
        for report_id in list(set(df_rls.New_Report_Id.values)):
            update_DataModelRoleAssignments(df_rls=df_rls, report_id=report_id)
        print("Step 6: 创建Row-Level-Security成功")
    except Exception as e:
        print("Step 6: 创建Row-Level-Security失败", e)

    print("已全部完成")

    refresh_plans = get_CacheRefreshPlans_for_all_reports()
    for refresh_id in refresh_plans.Id.values:
        refresh_dashboard(refresh_lan_id=refresh_id)

except Exception as e:
    print(e)
