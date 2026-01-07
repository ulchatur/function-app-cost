import datetime
import json
import requests
import os
import io
from openpyxl import Workbook
import azure.functions as func

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
SUBSCRIPTION_ID = os.environ["SUBSCRIPTION_ID"]

def get_access_token():
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

def get_previous_month_range():
    today = datetime.date.today()
    first_day_this_month = today.replace(day=1)
    last_day_prev_month = first_day_this_month - datetime.timedelta(days=1)
    first_day_prev_month = last_day_prev_month.replace(day=1)
    return first_day_prev_month.isoformat(), last_day_prev_month.isoformat()

def fetch_cost(token, start_date, end_date):
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

def generate_excel(cost, start_date, end_date):
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

def main(req: func.HttpRequest) -> func.HttpResponse:
    try:
        token = get_access_token()
        start_date, end_date = get_previous_month_range()
        cost_data = fetch_cost(token, start_date, end_date)
        excel_file = generate_excel(cost_data, start_date, end_date)

        filename = f"azure_cost_{start_date}_to_{end_date}.xlsx"

        return func.HttpResponse(
            body=excel_file.read(),
            status_code=200,
            headers={
                "Content-Disposition": f"attachment; filename={filename}",
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            }
        )

    except Exception as e:
        return func.HttpResponse(
            body=json.dumps({"error": str(e)}),
            status_code=500,
            mimetype="application/json"
        )
