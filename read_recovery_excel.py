from pbrs_api_function import *

recovery_file_path = None


def get_recovery_excel_path():
    cwd = os.getcwd()
    latest_file = get_latest_file(pbrs_excel_info_path)
    latest_file_absolut = os.path.join(cwd, latest_file)
    return latest_file_absolut


if recovery_file_path is None:
    recovery_file_path = get_recovery_excel_path()
    print("recovery_file_path", recovery_file_path)
else:
    print("recovery_file_path", recovery_file_path)

try:
    folders = pd.read_excel(recovery_file_path, sheet_name="Folders")
    powerbireports = pd.read_excel(recovery_file_path, sheet_name="PowerBIReports")
    schedules = pd.read_excel(recovery_file_path, sheet_name="Schedules")
    refresh_plans = pd.read_excel(recovery_file_path, sheet_name="CacheRefreshPlans")
    policies_user = pd.read_excel(recovery_file_path, sheet_name="Policies")
    DataModelRoles = pd.read_excel(recovery_file_path, sheet_name="DataModelRoles")
    DataModelRoleAssignments = pd.read_excel(recovery_file_path, sheet_name="DataModelRoleAssignments")
except Exception as e:
    print("read {} fails".format(recovery_file_path))
    print(e)
