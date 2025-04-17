# app.py
import io
import json
from datetime import datetime, timedelta, timezone
import pandas as pd
import numpy as np
import pytz
import requests
import time
from flask import Flask, request, render_template, send_file, flash, redirect, url_for

# --- Global configuration variables ---
api_key = "pk_3326657_EOM3G6Z3CKH2W61H8NOL5T7AGO9D7LNN"
team_id = "3314662"
__version__ = "v3.0.2"
__date__ = "11th December 2024"
__auth__ = api_key

# (Include here the month and project lists and other globals if needed)
columns_to_check = [
    "Course", "Product", "Proj-Common-Activity", "Proj-Outside-Office",
    "Management-Project", "Technology-Project", "Linguistic-Project",
    "MMedia-Project", "Project-CST", "Sales-Mktg-Project", "Project-ELA",
    "Proj-KidsPersona", "Project-Finance", "Website", "SFH-Admin-Project",
    "Admin-Project", "Linguistic-Activity"
]
project_columns = [
    "Proj-Common-Activity", "Proj-Outside-Office",
    "Management-Project", "Technology-Project", "Linguistic-Project",
    "MMedia-Project", "Project-CST", "Sales-Mktg-Project", "Project-ELA",
    "Proj-KidsPersona", "Project-Finance", "SFH-Admin-Project",
    "Admin-Project", "Linguistic-Activity"
]
ist_timezone = pytz.timezone('Asia/Kolkata')

app = Flask(__name__)
app.secret_key = "your_secret_key_here"   # Needed for flash messages

# --- Helper Functions (copy parts of your original functions) ---

def is_nan(value):
    return value == 'nan' or (isinstance(value, float) and np.isnan(value))

def convert_milliseconds_to_hours_minutes(milliseconds):
    seconds = milliseconds / 1000
    minutes = seconds // 60
    hours = minutes // 60
    minutes = minutes % 60
    return (int(hours), int(minutes))

def memberInfo():
    url = "https://api.clickup.com/api/v2/team"
    headers = {"Authorization": __auth__}
    response = requests.get(url, headers=headers)
    data = response.json()
    members_dict = {}
    for team in data.get('teams', []):
        for member in team.get('members', []):
            member_id = member['user']['id']
            member_username = member['user']['username']
            members_dict[member_id] = member_username
    # Map last 4 digits of username to member id
    members_dict = {value[-4:]: key for key, value in members_dict.items() if value is not None}
    return members_dict

def generate_timesheet(employee_id, start_date_str, end_date_str, open_google_sheet):
    """
    This is your main processing function. It replicates the behavior of get_selected_dates from your original code.
    It processes the data, fetches from the API and returns the Excel file as bytes.
    """
    # Convert the date strings from the form into date objects:
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()
    key = employee_id.upper()
    start_date_fmt = start_date.strftime("%b %d")
    end_date_fmt = end_date.strftime("%b %d")
    year_str = str(start_date.year)
    filename = f"{key}_{start_date_fmt}_to_{end_date_fmt}_{year_str}.xlsx"
    
    start_time_process = time.time()
    
    members_dict = memberInfo()
    if key not in members_dict:
        raise Exception("Invalid Employee ID. Please check and try again.")
    employee_key = members_dict[key]

    # Convert dates to timestamps
    start_datetime = datetime.combine(start_date, datetime.min.time())
    start_timestamp = int(start_datetime.replace(tzinfo=timezone.utc).timestamp())
    end_datetime = datetime.combine(end_date, datetime.min.time())
    end_timestamp = int(end_datetime.replace(tzinfo=timezone.utc).timestamp())

    url = f"https://api.clickup.com/api/v2/team/{team_id}/time_entries"
    query = {
        "start_date": str(int(start_timestamp - 19800) * 1000),
        "end_date": str(int((end_timestamp + 86399) * 1000) - 19800000),
        "assignee": employee_key,
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": __auth__
    }
    response = requests.get(url, headers=headers, params=query)
    data = response.json()
    if 'data' not in data or not data['data']:
        raise Exception("There are no entries in this Date Range. Please change Date Range or update entries in ClickUp.")
    
    # Process data into a DataFrame
    task_names, task_ids, task_status, durations, dates, days = [], [], [], [], [], []
    for entry in data['data']:
        try:
            task_names.append(entry['task']['name'])
            task_ids.append(entry['task']['id'])
            task_status.append(entry['task']['status']['status'])
        except Exception:
            task_names.append('0')
            task_ids.append('0')
            task_status.append('0')
        durations.append(int(entry['duration']))
        start_time = int(entry['start']) // 1000
        date_val = pd.Timestamp(start_time, unit='s').date()
        dates.append(date_val)
        localized_start_datetime = pytz.utc.localize(datetime.utcfromtimestamp(start_time))
        day = localized_start_datetime.astimezone(ist_timezone).strftime('%A')
        days.append(day)

    df = pd.DataFrame({
        'Task Name': task_names,
        'Task ID': task_ids,
        'Task Status': task_status,
        'Duration': durations,
        'Date': dates,
        'Day': days
    })

    unique_task_ids = df['Task ID'].unique()
    new_df = pd.DataFrame({'Task ID': unique_task_ids})
    days_of_week = ['Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    for day in days_of_week:
        new_df[day] = 0

    for task in unique_task_ids:
        task_entries = df[df['Task ID'] == task]
        grouped_entries = task_entries.groupby(['Day']).sum(numeric_only=True)
        for day in days_of_week:
            if day in grouped_entries.index:
                new_df.loc[new_df['Task ID'] == task, day] = grouped_entries.loc[day]['Duration']

    df_h = pd.merge(df, new_df, on='Task ID')
    df_h.drop_duplicates(subset='Task ID', inplace=True)
    df_h[days_of_week] = df_h[days_of_week].apply(lambda x: x / 3600000).round(2)
    df_h = df_h.drop(['Duration', 'Date', 'Day'], axis=1)

    headers_api = {"Authorization": __auth__}
    for task_id in df_h['Task ID'].unique():
        task_url = f"https://api.clickup.com/api/v2/task/{task_id}"
        response = requests.get(task_url, headers=headers_api)
        tasks = response.json()
        hrs_mins = convert_milliseconds_to_hours_minutes(tasks.get('time_spent', 0))
        df_h.loc[df_h['Task ID'] == task_id,
                 'Total Time tracked for this task till now (hrs)'] = f"{hrs_mins[0]}h {hrs_mins[1]}m"
        try:
            # Safely fetch custom fields
            custom_fields = tasks.get("custom_fields", [])
            for custom_field in custom_fields:
                if 'value' in custom_field and custom_field['type'] == 'drop_down':
                    df_h.loc[df_h['Task ID'] == task_id, custom_field['name']] = custom_field['type_config']['options'][custom_field['value']]['name']
        except Exception as e:
            # Log the error details
            print(f"Error processing custom fields for task {task_id}: {e}")
            print("Task details:", json.dumps(tasks, indent=2))
            continue

    # Your additional filtering and summarizing logic goes here…
    df_h['Total Tracked this week in this task'] = df_h[days_of_week].sum(axis=1)
    totals = df_h[days_of_week].sum(axis=0)
    df_h = pd.concat([df_h, totals.to_frame().T], ignore_index=True)
    weekly_total = df_h.iloc[-1].sum()
    df_h.at[df_h.index[-1], 'Task Status'] = 'Daily Totals ->'
    empty_row = pd.Series([np.nan] * len(df_h.columns), index=df_h.columns)
    df_h = pd.concat([df_h, empty_row.to_frame().T], ignore_index=True)
    df_h.iloc[-1, 5] = weekly_total

    days_diff = (end_date - start_date).days + 1
    if days_diff <= 7:
        df_h.iloc[:, 3] = df_h.iloc[:, 3].astype(object)
        df_h.iloc[-1, 3] = "Week's total ="
        week_number = end_date.isocalendar()[1]
        df_h.at[df_h.index[-1], 'Task Name'] = f'Week #{week_number} - {start_date_fmt}, {year_str} - {end_date_fmt}, {year_str}'
    else:
        df_h.iloc[-1, 3] = 'Total Hours Tracked ='
        df_h.at[df_h.index[-1], 'Task Name'] = f'{start_date_fmt}, {year_str} - {end_date_fmt}, {year_str}'

    df_h.insert(10, 'Total Tracked this week in this task', df_h.pop('Total Tracked this week in this task'))

    # Write to Excel file in memory
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df_h.to_excel(writer, sheet_name='Sheet1', index=False)
    worksheet = writer.sheets['Sheet1']
    for row_num, value in enumerate(df_h['Task ID'], start=1):
        if pd.isna(value):
            break
        url = f'https://app.clickup.com/t/{value}'
        worksheet.write_url(row_num, df_h.columns.get_loc('Task ID'), url, string=value)
    writer.close()
    processed_data = output.getvalue()

    elapsed = time.time() - start_time_process
    print(f"Processing time: {elapsed:.2f} seconds")

    # Return the filename and the byte stream
    return filename, processed_data

# --- Flask Routes ---

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        employee_id = request.form.get("employee_id")
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")
        open_google_sheet = request.form.get("open_google_sheet") == "on"
        try:
            filename, file_data = generate_timesheet(employee_id, start_date, end_date, open_google_sheet)
            # Return the file as an attachment
            return send_file(
                io.BytesIO(file_data),
                as_attachment=True,
                download_name=filename,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            flash(str(e), "error")
            return redirect(url_for("index"))
    return render_template("index.html", version=__version__, build_date=__date__)

if __name__ == "__main__":
    app.run(debug=True)
