from pbrs_api_function import *

refresh_plans = get_CacheRefreshPlans_for_all_reports()
for refresh_id in refresh_plans.Id.values:
    refresh_dashboard(refresh_lan_id=refresh_id)
