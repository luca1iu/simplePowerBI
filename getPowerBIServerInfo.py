from pbrs_api_function import *

# create a pandas Excel writer object
writer = pd.ExcelWriter(r'..\PowerBiReportSever_Info.xlsx')

folders = get_df_from_pbrs(Folders)
powerbireports = get_df_from_pbrs(PowerBIReports)
schedules = get_df_from_pbrs(Schedules)
refresh_plans = get_CacheRefreshPlans_for_all_reports()
policies_user = get_Policies_for_all_reports()
DataModelRoles = get_DataModelRoles_for_all_reports()
DataModelRoleAssignments = get_DataModelRoleAssignments_for_all_reports()

# write each DataFrame to the writer object
folders.to_excel(writer, sheet_name='Folders', index=False)
powerbireports.to_excel(writer, sheet_name='PowerBIReports', index=False)
schedules.to_excel(writer, sheet_name='Schedules', index=False)
refresh_plans.to_excel(writer, sheet_name='CacheRefreshPlans', index=False)
policies_user.to_excel(writer, sheet_name='Policies', index=False)
DataModelRoles.to_excel(writer, sheet_name='DataModelRoles', index=False)
DataModelRoleAssignments.to_excel(writer, sheet_name='DataModelRoleAssignments', index=False)

# save the Excel file
writer.save()
print("保存成功")
