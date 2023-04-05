import requests
import base64
import json
import pandas as pd
import openpyxl
import smtplib
from datetime import datetime
import os
from typing import List, Dict
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# Azure DevOps API settings
organization = "DevOpsCelebal"
project = "Chitransh"
api_version = "7.1-preview.3"
personal_access_token = ""

# Email settings
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "you@gmail.com"
receiver_emails = ["person1@example.com", "person2@example.com"]
email_password = "your_email_password"


def get_epics():
    url = (
        f"https://dev.azure.com/{organization}/{project}/_apis/wit/wiql?api-version=6.0"
    )
    encoded_pat = base64.b64encode((":" + personal_access_token).encode()).decode()
    headers = {"Authorization": "Basic " + encoded_pat}
    query = {
        "query": "SELECT [System.Id] FROM workitems WHERE [System.WorkItemType] = 'Epic'"
    }
    response = requests.post(url, headers=headers, json=query)

    if response.status_code != 200:
        print(f"Error: API request failed with status code {response.status_code}")
        print(response.text)
        return []

    data = response.json()
    work_items = [work_item["id"] for work_item in data["workItems"]]
    return work_items


def get_work_items(epic_ids):
    work_items = []
    encoded_pat = base64.b64encode((":" + personal_access_token).encode()).decode()
    headers = {"Authorization": "Basic " + encoded_pat}

    for epic_id in epic_ids:
        url = f"https://dev.azure.com/{organization}/{project}/_apis/wit/workitems/{epic_id}?api-version={api_version}&$expand=all"
        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            print(f"Error: API request failed with status code {response.status_code}")
            print(response.text)
            continue

        data = response.json()
        work_items.append(data)

    return work_items


# Function to create an Excel file with the work items data
def create_excel(work_items):
    work_items_data = []

    for work_item_id in work_items:
        url = f"https://dev.azure.com/{organization}/{project}/_apis/wit/workItems/{work_item_id}?api-version={api_version}"
        headers = {"Authorization": "Basic " + personal_access_token}
        response = requests.get(url, headers=headers)
        data = response.json()

        work_items_data.append(
            {
                "ID": data["id"],
                "Title": data["fields"]["System.Title"],
                "Assigned To": data["fields"]["System.AssignedTo"]["displayName"],
                "State": data["fields"]["System.State"],
                "Area Path": data["fields"]["System.AreaPath"],
                "Iteration Path": data["fields"]["System.IterationPath"],
            }
        )

    df = pd.DataFrame(work_items_data)
    excel_file = "Epic_Work_Items.xlsx"
    df.to_excel(excel_file, index=False)

    return excel_file


def write_epics_to_excel(epic_list: List[Dict]) -> None:
    workbook = openpyxl.Workbook()
    del workbook["Sheet"]

    # Create sheets for each work item state
    todo_sheet = workbook.create_sheet("To Do")
    doing_sheet = workbook.create_sheet("Doing")
    done_sheet = workbook.create_sheet("Done")

    state_sheets = {
        "To Do": todo_sheet,
        "Doing": doing_sheet,
        "Done": done_sheet,
    }

    # Formatting
    header_font = openpyxl.styles.Font(bold=True)
    header_fill = openpyxl.styles.PatternFill(
        start_color="FFFF00", end_color="FFFF00", fill_type="solid"
    )
    header_alignment = openpyxl.styles.Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

    # Function to write header row to a sheet
    def write_header(sheet):
        headers = [
            "ID",
            "Title",
            "State",
            "Priority",
            "Start Date",
            "Target Date",
            "Description",
        ]
        for col, header in enumerate(headers, 1):
            cell = sheet.cell(row=1, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment

    # Write header rows to each sheet
    for sheet in state_sheets.values():
        write_header(sheet)

    # Convert the RevisedDate string to a datetime object if it's a string
    for epic in epic_list:
        if isinstance(epic["fields"]["System.RevisedDate"], str):
            epic["fields"]["System.RevisedDate"] = datetime.fromisoformat(
                epic["fields"]["System.RevisedDate"].rstrip("Z")
            )

    # Sort the epics by RevisedDate in descending order
    sorted_epics = sorted(
        epic_list, key=lambda x: x["fields"]["System.RevisedDate"], reverse=True
    )

    # Write data rows to respective sheets based on work item state
    row_counters = {"To Do": 2, "Doing": 2, "Done": 2}

    for epic in sorted_epics:
        state = epic["fields"]["System.State"]
        sheet = state_sheets[state]
        row = row_counters[state]

        sheet.cell(row=row, column=1, value=epic.get("id"))
        sheet.cell(row=row, column=2, value=epic.get("fields", {}).get("System.Title"))
        sheet.cell(row=row, column=3, value=epic.get("fields", {}).get("System.State"))
        sheet.cell(
            row=row,
            column=4,
            value=epic.get("fields", {}).get("Microsoft.VSTS.Common.Priority", ""),
        )
        sheet.cell(
            row=row,
            column=5,
            value=epic.get("fields", {}).get("Microsoft.VSTS.Scheduling.StartDate", ""),
        )
        sheet.cell(
            row=row,
            column=6,
            value=epic.get("fields", {}).get(
                "Microsoft.VSTS.Scheduling.TargetDate", ""
            ),
        )
        sheet.cell(
            row=row,
            column=7,
            value=epic.get("fields", {}).get("System.Description", ""),
        )

        row_counters[state] += 1

    # Adjust column widths for each sheet
    for sheet in state_sheets.values():
        for col in range(1, 8):
            sheet.column_dimensions[
                openpyxl.utils.get_column_letter(col)
            ].auto_size = True

    # Save the Excel file
    workbook.save("epics.xlsx")


# Function to send an email with the Excel attachment
def send_email(excel_file, sender_email, receiver_emails, email_password):
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = ", ".join(receiver_emails)
    msg["Subject"] = "Epic Work Items Report"

    body = "Please find the attached Epic Work Items Report."
    msg.attach(MIMEText(body, "plain"))

    attachment = open(excel_file, "rb")
    base = MIMEBase("application", "octet-stream")
    base.set_payload((attachment).read())
    encoders.encode_base64(base)
    base.add_header("Content-Disposition", "attachment; filename=%s" % excel_file)
    msg.attach(base)

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, email_password)
    text = msg.as_string()
    server.sendmail(sender_email, receiver_emails, text)
    server.quit()


# Main script
if __name__ == "__main__":
    epics = get_epics()
    work_items = get_work_items(epics)
    write_epics_to_excel(epic_list=work_items)
    # excel_file = create_excel(work_items)
    # send_email(excel_file, sender_email, receiver_emails, email_password)

    # # Delete the Excel file
    # os.remove(excel_file)
