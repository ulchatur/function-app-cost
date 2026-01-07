import datetime
import json
import requests
import os
import io
import logging
from openpyxl import Workbook
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

def get_previous_month_range():
    today = datetime.date.today()
    first_day_this_month = today.replace(day=1)
    last_day_prev_month = first_day_this_month - datetime.timedelta(days=1)
    first_day_prev_month = last_day_prev_month.replace(day=1)
    return first_day_prev_month.isoformat(), last_day_prev_month.isoformat()

def fetch_cost(token, start_date, end_date):
    try:
        SUBSCRIPTION_ID = os.environ.get("SUBSCRIPTION_ID")
        
        if not SUBSCRIPTION_ID:
            raise ValueError("Missing environment variable: SUBSCRIPTION_ID")
        
        url = f"https://management.azure.com/subscriptions/{SUBSCRIPTION_ID}/providers/Microsoft.CostManagement/query?api-version=2023-03-01"

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
        logger.error(f"Error fetching cost: {str(e)}")
        raise

def generate_excel(cost, start_date, end_date):
    try:
        SUBSCRIPTION_ID = os.environ.get("SUBSCRIPTION_ID")
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Azure Cost Report"

        ws.append([
            "Subscription ID",
            "From Date",
            "To Date",
            "Total Cost"
        ])

        total_cost = cost["properties"]["rows"][0][0]

        ws.append([
            SUBSCRIPTION_ID,
            start_date,
            end_date,
            total_cost
        ])

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
        required_vars = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "SUBSCRIPTION_ID"]
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
        
        logger.info("Getting date range...")
        start_date, end_date = get_previous_month_range()
        
        logger.info(f"Fetching cost data from {start_date} to {end_date}...")
        cost_data = fetch_cost(token, start_date, end_date)
        
        logger.info("Generating Excel file...")
        excel_file = generate_excel(cost_data, start_date, end_date)

        filename = f"azure_cost_{start_date}_to_{end_date}.xlsx"

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