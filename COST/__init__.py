# import datetime
# import json
# import requests
# import os
# import io
# import logging
# from openpyxl import Workbook
# from openpyxl.styles import Font, PatternFill, Alignment
# import azure.functions as func

# # Setup logging
# logger = logging.getLogger(__name__)

# def get_access_token():
#     try:
#         TENANT_ID = os.environ.get("TENANT_ID")
#         CLIENT_ID = os.environ.get("CLIENT_ID")
#         CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
        
#         if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
#             raise ValueError("Missing environment variables: TENANT_ID, CLIENT_ID, or CLIENT_SECRET")
        
#         url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/token"
#         payload = {
#             "grant_type": "client_credentials",
#             "client_id": CLIENT_ID,
#             "client_secret": CLIENT_SECRET,
#             "resource": "https://management.azure.com/"
#         }
#         response = requests.post(url, data=payload)
#         response.raise_for_status()
#         return response.json()["access_token"]
#     except Exception as e:
#         logger.error(f"Error getting access token: {str(e)}")
#         raise

# def get_previous_month_range():
#     today = datetime.date.today()
#     first_day_this_month = today.replace(day=1)
#     last_day_prev_month = first_day_this_month - datetime.timedelta(days=1)
#     first_day_prev_month = last_day_prev_month.replace(day=1)
#     return first_day_prev_month.isoformat(), last_day_prev_month.isoformat()

# def get_all_subscriptions(token):
#     """Fetch all subscriptions accessible to the service principal"""
#     try:
#         url = "https://management.azure.com/subscriptions?api-version=2020-01-01"
#         headers = {
#             "Authorization": f"Bearer {token}",
#             "Content-Type": "application/json"
#         }
        
#         response = requests.get(url, headers=headers)
#         response.raise_for_status()
        
#         subscriptions = response.json().get("value", [])
#         logger.info(f"Found {len(subscriptions)} subscriptions")
        
#         return subscriptions
#     except Exception as e:
#         logger.error(f"Error fetching subscriptions: {str(e)}")
#         raise

# def fetch_cost_for_subscription(token, subscription_id, start_date, end_date):
#     """Fetch cost data for a specific subscription"""
#     try:
#         url = f"https://management.azure.com/subscriptions/{subscription_id}/providers/Microsoft.CostManagement/query?api-version=2023-03-01"

#         headers = {
#             "Authorization": f"Bearer {token}",
#             "Content-Type": "application/json"
#         }

#         body = {
#             "type": "ActualCost",
#             "timeframe": "Custom",
#             "timePeriod": {
#                 "from": start_date,
#                 "to": end_date
#             },
#             "dataset": {
#                 "granularity": "None",
#                 "aggregation": {
#                     "totalCost": {
#                         "name": "Cost",
#                         "function": "Sum"
#                     }
#                 }
#             }
#         }

#         response = requests.post(url, headers=headers, json=body)
#         response.raise_for_status()
#         return response.json()
#     except Exception as e:
#         logger.error(f"Error fetching cost for subscription {subscription_id}: {str(e)}")
#         # Return empty structure if cost fetch fails
#         return {"properties": {"rows": [], "columns": []}}

# def generate_excel(all_costs_data, start_date, end_date):
#     """Generate Excel with all subscriptions cost data"""
#     try:
#         wb = Workbook()
#         ws = wb.active
#         ws.title = "Azure Cost Report"

#         # Header styling
#         header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
#         header_font = Font(bold=True, color="FFFFFF")
        
#         # Add headers
#         headers = ["Subscription Name", "Subscription ID", "From Date", "To Date", "Total Cost (USD)", "Status"]
#         ws.append(headers)
        
#         # Style header row
#         for cell in ws[1]:
#             cell.fill = header_fill
#             cell.font = header_font
#             cell.alignment = Alignment(horizontal="center", vertical="center")

#         total_cost_all = 0.0
        
#         # Add data for each subscription
#         for sub_data in all_costs_data:
#             subscription_name = sub_data["subscription_name"]
#             subscription_id = sub_data["subscription_id"]
#             cost_data = sub_data["cost_data"]
            
#             # Extract cost
#             rows = cost_data.get("properties", {}).get("rows", [])
            
#             if not rows or len(rows) == 0:
#                 total_cost = 0.0
#                 status = "No usage data"
#             else:
#                 total_cost = float(rows[0][0]) if len(rows[0]) > 0 else 0.0
#                 status = "Active"
            
#             total_cost_all += total_cost
            
#             ws.append([
#                 subscription_name,
#                 subscription_id,
#                 start_date,
#                 end_date,
#                 round(total_cost, 2),
#                 status
#             ])

#         # Add total row
#         ws.append([])
#         total_row = ws.max_row
#         ws.append(["TOTAL", "", "", "", round(total_cost_all, 2), ""])
        
#         # Style total row
#         for cell in ws[total_row]:
#             cell.font = Font(bold=True)
#             cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")

#         # Adjust column widths
#         ws.column_dimensions['A'].width = 35
#         ws.column_dimensions['B'].width = 40
#         ws.column_dimensions['C'].width = 15
#         ws.column_dimensions['D'].width = 15
#         ws.column_dimensions['E'].width = 20
#         ws.column_dimensions['F'].width = 15

#         # Add summary info
#         ws.append([])
#         ws.append([f"Report Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
#         ws.append([f"Total Subscriptions: {len(all_costs_data)}"])

#         file_stream = io.BytesIO()
#         wb.save(file_stream)
#         file_stream.seek(0)
#         return file_stream
#     except Exception as e:
#         logger.error(f"Error generating excel: {str(e)}")
#         raise

# def main(req: func.HttpRequest) -> func.HttpResponse:
#     logger.info('Python HTTP trigger function processed a request.')
    
#     try:
#         # Check environment variables
#         required_vars = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET"]
#         missing_vars = [var for var in required_vars if not os.environ.get(var)]
        
#         if missing_vars:
#             error_msg = f"Missing environment variables: {', '.join(missing_vars)}"
#             logger.error(error_msg)
#             return func.HttpResponse(
#                 body=json.dumps({"error": error_msg}),
#                 status_code=500,
#                 mimetype="application/json"
#             )
        
#         logger.info("Getting access token...")
#         token = get_access_token()
        
#         logger.info("Getting date range...")
#         start_date, end_date = get_previous_month_range()
#         logger.info(f"Date range: {start_date} to {end_date}")
        
#         logger.info("Fetching all subscriptions...")
#         subscriptions = get_all_subscriptions(token)
        
#         if not subscriptions:
#             return func.HttpResponse(
#                 body=json.dumps({"error": "No subscriptions found"}),
#                 status_code=404,
#                 mimetype="application/json"
#             )
        
#         logger.info(f"Processing {len(subscriptions)} subscriptions...")
        
#         # Fetch cost for each subscription
#         all_costs_data = []
#         for subscription in subscriptions:
#             sub_id = subscription.get("subscriptionId")
#             sub_name = subscription.get("displayName", "Unknown")
            
#             logger.info(f"Fetching cost for: {sub_name} ({sub_id})")
            
#             cost_data = fetch_cost_for_subscription(token, sub_id, start_date, end_date)
            
#             all_costs_data.append({
#                 "subscription_id": sub_id,
#                 "subscription_name": sub_name,
#                 "cost_data": cost_data
#             })
        
#         logger.info("Generating Excel file...")
#         excel_file = generate_excel(all_costs_data, start_date, end_date)

#         filename = f"azure_all_subscriptions_cost_{start_date}_to_{end_date}.xlsx"

#         logger.info("Returning Excel file...")
#         return func.HttpResponse(
#             body=excel_file.read(),
#             status_code=200,
#             headers={
#                 "Content-Disposition": f"attachment; filename={filename}",
#                 "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             }
#         )

#     except Exception as e:
#         error_msg = f"Error: {str(e)}"
#         logger.error(error_msg, exc_info=True)
#         return func.HttpResponse(
#             body=json.dumps({
#                 "error": str(e),
#                 "type": type(e).__name__
#             }),
#             status_code=500,
#             mimetype="application/json"
#         )





import datetime
import json
import requests
import os
import io
import logging
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import azure.functions as func

# Setup logging
logger = logging.getLogger(__name__)

def get_access_token():
    try:
        TENANT_ID = os.environ.get("TENANT_ID")
        CLIENT_ID = os.environ.get("CLIENT_ID")
        CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
        
        if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
            raise ValueError("Missing environment variables: TENANT_ID, CLIENT_ID, or CLIENT_SECRET")
        
        url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/token"
        payload = {
            "grant_type": "client_credentials",
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "resource": "https://management.azure.com/"
        }
        response = requests.post(url, data=payload)
        response.raise_for_status()
        return response.json()["access_token"]
    except Exception as e:
        logger.error(f"Error getting access token: {str(e)}")
        raise

def get_month_range(months_back=1):
    """
    Get date range for a specific month in the past
    months_back=1 means previous month
    months_back=2 means 2 months ago, etc.
    """
    today = datetime.date.today()
    
    # Start from current month
    current_month = today.replace(day=1)
    
    # Go back the specified number of months
    target_month = current_month
    for _ in range(months_back):
        target_month = (target_month - datetime.timedelta(days=1)).replace(day=1)
    
    # Get last day of target month
    next_month = target_month.replace(day=28) + datetime.timedelta(days=4)
    last_day = (next_month - datetime.timedelta(days=next_month.day)).replace(day=target_month.day)
    
    # Find actual last day
    if target_month.month == 12:
        last_day = target_month.replace(day=31)
    else:
        next_month_first = target_month.replace(month=target_month.month + 1, day=1)
        last_day = next_month_first - datetime.timedelta(days=1)
    
    logger.info(f"Today: {today}")
    logger.info(f"Target month ({months_back} months back): {target_month} to {last_day}")
    
    return target_month.isoformat(), last_day.isoformat()

def get_previous_month_range():
    """Get previous month (1 month back)"""
    return get_month_range(1)

def get_current_month_range():
    """Get current month date range for testing"""
    today = datetime.date.today()
    first_day_this_month = today.replace(day=1)
    return first_day_this_month.isoformat(), today.isoformat()

def get_all_subscriptions(token):
    """Fetch all subscriptions accessible to the service principal"""
    try:
        url = "https://management.azure.com/subscriptions?api-version=2020-01-01"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        subscriptions = response.json().get("value", [])
        logger.info(f"Found {len(subscriptions)} subscriptions")
        
        return subscriptions
    except Exception as e:
        logger.error(f"Error fetching subscriptions: {str(e)}")
        raise

def fetch_cost_for_subscription(token, subscription_id, start_date, end_date):
    """Fetch cost data for a specific subscription"""
    try:
        url = f"https://management.azure.com/subscriptions/{subscription_id}/providers/Microsoft.CostManagement/query?api-version=2023-03-01"

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        body = {
            "type": "ActualCost",
            "timeframe": "Custom",
            "timePeriod": {
                "from": start_date,
                "to": end_date
            },
            "dataset": {
                "granularity": "None",
                "aggregation": {
                    "totalCost": {
                        "name": "Cost",
                        "function": "Sum"
                    }
                }
            }
        }

        response = requests.post(url, headers=headers, json=body)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        logger.error(f"Error fetching cost for subscription {subscription_id}: {str(e)}")
        # Return empty structure if cost fetch fails
        return {"properties": {"rows": [], "columns": []}}

def generate_excel(all_costs_data, start_date, end_date):
    """Generate Excel with all subscriptions cost data"""
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Azure Cost Report"

        # Header styling
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        # Add headers
        headers = ["Subscription Name", "Subscription ID", "From Date", "To Date", "Total Cost (USD)", "Status"]
        ws.append(headers)
        
        # Style header row
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        total_cost_all = 0.0
        
        # Add data for each subscription
        for sub_data in all_costs_data:
            subscription_name = sub_data["subscription_name"]
            subscription_id = sub_data["subscription_id"]
            cost_data = sub_data["cost_data"]
            
            # Extract cost
            rows = cost_data.get("properties", {}).get("rows", [])
            
            if not rows or len(rows) == 0:
                total_cost = 0.0
                status = "No usage data"
            else:
                total_cost = float(rows[0][0]) if len(rows[0]) > 0 else 0.0
                status = "Active"
            
            total_cost_all += total_cost
            
            ws.append([
                subscription_name,
                subscription_id,
                start_date,
                end_date,
                round(total_cost, 2),
                status
            ])

        # Add total row
        ws.append([])
        total_row = ws.max_row
        ws.append(["TOTAL", "", "", "", round(total_cost_all, 2), ""])
        
        # Style total row
        for cell in ws[total_row]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")

        # Adjust column widths
        ws.column_dimensions['A'].width = 35
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 20
        ws.column_dimensions['F'].width = 15

        # Add summary info
        ws.append([])
        ws.append([f"Report Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
        ws.append([f"Total Subscriptions: {len(all_costs_data)}"])

        file_stream = io.BytesIO()
        wb.save(file_stream)
        file_stream.seek(0)
        return file_stream
    except Exception as e:
        logger.error(f"Error generating excel: {str(e)}")
        raise

def main(req: func.HttpRequest) -> func.HttpResponse:
    logger.info('Python HTTP trigger function processed a request.')
    
    try:
        # Check environment variables
        required_vars = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET"]
        missing_vars = [var for var in required_vars if not os.environ.get(var)]
        
        if missing_vars:
            error_msg = f"Missing environment variables: {', '.join(missing_vars)}"
            logger.error(error_msg)
            return func.HttpResponse(
                body=json.dumps({"error": error_msg}),
                status_code=500,
                mimetype="application/json"
            )
        
        logger.info("Getting access token...")
        token = get_access_token()
        
        # Check query parameter for date range
        date_range = req.params.get('range', 'previous')  # 'previous' or 'current'
        
        logger.info("Getting date range...")
        if date_range == 'current':
            start_date, end_date = get_current_month_range()
            logger.info(f"Fetching CURRENT month cost: {start_date} to {end_date}")
        else:
            start_date, end_date = get_previous_month_range()
            logger.info(f"Fetching PREVIOUS month cost: {start_date} to {end_date}")
        
        logger.info("Fetching all subscriptions...")
        subscriptions = get_all_subscriptions(token)
        
        if not subscriptions:
            return func.HttpResponse(
                body=json.dumps({"error": "No subscriptions found"}),
                status_code=404,
                mimetype="application/json"
            )
        
        logger.info(f"Processing {len(subscriptions)} subscriptions...")
        
        # Fetch cost for each subscription
        all_costs_data = []
        for subscription in subscriptions:
            sub_id = subscription.get("subscriptionId")
            sub_name = subscription.get("displayName", "Unknown")
            
            logger.info(f"Fetching cost for: {sub_name} ({sub_id})")
            
            cost_data = fetch_cost_for_subscription(token, sub_id, start_date, end_date)
            
            all_costs_data.append({
                "subscription_id": sub_id,
                "subscription_name": sub_name,
                "cost_data": cost_data
            })
        
        logger.info("Generating Excel file...")
        excel_file = generate_excel(all_costs_data, start_date, end_date)

        filename = f"azure_all_subscriptions_cost_{start_date}_to_{end_date}.xlsx"

        logger.info("Returning Excel file...")
        return func.HttpResponse(
            body=excel_file.read(),
            status_code=200,
            headers={
                "Content-Disposition": f"attachment; filename={filename}",
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
        )

    except Exception as e:
        error_msg = f"Error: {str(e)}"
        logger.error(error_msg, exc_info=True)
        return func.HttpResponse(
            body=json.dumps({
                "error": str(e),
                "type": type(e).__name__
            }),
            status_code=500,
            mimetype="application/json"
        )