import pandas as pd
from tkinter import messagebox, filedialog
import requests
import json
import csv
from ttkbootstrap.constants import *
from ttkbootstrap.tooltip import ToolTip
import ttkbootstrap as tb
import os

base_dir = os.path.dirname(__file__)
CONFIG_FILE = os.path.join(base_dir, './config.json')


def process_excel_file():
    try:
        with open(CONFIG_FILE, "r") as file:
            config_data = json.load(file)
    except FileNotFoundError:
        config_data = {"username": "", "api_token": "", "output_file": ""}

    try:
        sprint_data_file_name = 'sprint_data.csv'
        sprint_data_file_path = os.path.join(os.path.dirname(__file__), sprint_data_file_name)

        # Read the existing CSV file
        df = pd.read_csv(sprint_data_file_path)

        # Determine the 'Department' based on the 'Reporter' column
        conditions = [
            ((df['Reporter'] == 'Adriana Novo') | (df['Reporter'] == 'Thomas Agius'), 'Content'),
            ((df['Reporter'] == 'Christian Haensel') | (df['Reporter'] == 'Aleksandar Manok') | (df['Reporter'] == 'Marta Crovetto'), 'SEO'),
            ((df['Reporter'] == 'Tiberiu Petcu') | (df['Reporter'] == 'Mojtaba Darvishi') | (df['Reporter'] == 'Rodrigo Isidro') | (df['Reporter'] == 'Julio Pazini') | (df['Reporter'] == 'Poliana Rufatto') | (df['Reporter'] == 'Anton Micallef'), 'Tech'),
        ]

        # Create a new 'Department' column based on the conditions
        df['Department'] = [next((condition[1] for condition in conditions if condition[0].loc[i]), '') for i in range(len(df))]

        # Add a new 'Theme' column with the word 'Kaffiliate'
        df['Theme'] = 'Kaffiliate'

        # Set 'Theme' to 'KCS' if 'KCS' exists in 'Summary'
        df.loc[df['Summary'].str.contains('KCS'), 'Theme'] = 'KCS'

        # Select and reorder columns
        selected_columns = ['Issue key', 'Summary', 'Reporter', 'Department', 'Theme']
        df = df[selected_columns]

        sprint_id = sprint_entry.get()
        output_file_name = f'sprint_{sprint_id}.xlsx'
        output_file_path = os.path.join(config_data["output_file"], output_file_name)

        # Create a new Excel workbook and add a worksheet
        with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
            create_excel_file(df, writer)

        result_label.config(text=f"File saved to \n {output_file_path}", font=('Helvetica', 14), bootstyle='success')
        delete_file(sprint_data_file_path)
    except FileNotFoundError:
        result_label.config(text="File not found.", font=('Helvetica', 14), bootstyle='danger')
    except Exception as e:
        result_label.config(text=f"An error occurred: {e}", font=('Helvetica', 14), bootstyle='danger')


def create_excel_file(df, writer):
    df.to_excel(writer, index=False, sheet_name='Sheet1')

    # Access the workbook and the worksheet
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # Get the 'Issue key' column index
    issue_key_col = df.columns.get_loc('Issue key')

    # Define a cell format for hyperlinks
    url_format = workbook.add_format({'color': 'blue', 'underline': 1})

    # Add hyperlinks to 'Issue key' column
    for i, value in enumerate(df['Issue key']):
        url = f'https://kaferocks.atlassian.net/browse/{value}'  # Replace with the actual URL format
        worksheet.write_url(i + 1, issue_key_col, url, url_format, string=value)

    # Get the 'Department' column index
    content_col = df.columns.get_loc('Department')

    # Define a fill color for 'Content' cells (light green)
    content_format = workbook.add_format({'bg_color': '#92C47C'})

    # Apply conditional formatting to color 'Content' cells
    worksheet.conditional_format(1, content_col, len(df), content_col, {'type': 'text', 'criteria': 'containing', 'value': 'Content', 'format': content_format})

    # Adjust cell sizes to make text visible
    for col_num, col_width in enumerate(df.columns):
        # Set the width to the max length of the column content
        max_len = max(df[col_width].astype(str).apply(len).max(), len(col_width))
        worksheet.set_column(col_num, col_num, max_len + 2)  # Adding a little extra space


def get_tickets():
    # Load existing configuration from JSON file
    try:
        with open(CONFIG_FILE, "r") as file:
            config_data = json.load(file)
    except FileNotFoundError:
        config_data = {"username": "", "api_token": "", "output_file": ""}

    # Replace with your Jira details
    jira_url = 'https://kaferocks.atlassian.net'
    sprint_id = sprint_entry.get()
    api_url = f'{jira_url}/rest/agile/1.0/sprint/{sprint_id}/issue'

    username = config_data['username']
    api_token = config_data['api_token']

    # Make the API request
    response = requests.get(api_url, auth=(username, api_token))

    # Check for a successful response (status code 200)
    if response.status_code == 200:
        filter_excel_file(response)
    else:
        result_label.config(text="Empty fields..")
        print(f'Error: {response.status_code}')
        print(response.text)


def filter_excel_file(response):
    issues = response.json()['issues']

    # Filter issues based on status
    filtered_issues = [issue for issue in issues if
                    issue['fields']['status']['name'] in ['In Staging', 'Approved by Tech', 'In Progress']]

    file_path = os.path.join(os.path.dirname(__file__), 'sprint_data.csv')

    # Save filtered issues to a CSV file
    with open(file_path, 'w', newline='') as csvfile:
        fieldnames = ['Issue key', 'Summary', 'Reporter', 'Status']  # Adjust as needed
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        writer.writeheader()
        for issue in filtered_issues:
            writer.writerow({
                'Issue key': issue['key'],
                'Summary': issue['fields']['summary'],
                'Reporter': issue['fields']['reporter']['displayName'],
                'Status': issue['fields']['status']['name']
                # Add more fields as needed
            })

    print('Filtered data exported successfully.')
    result_label.config(text="Tickets retrieved!")


def process_sprint():
    get_tickets()
    process_excel_file()


def open_config_window():
    global output_entry
    config_window = tb.Toplevel(root)
    config_window.title("Config")
    config_window.geometry("400x350")

    # Load existing configuration from JSON file
    try:
        with open(CONFIG_FILE, "r") as file:
            config_data = json.load(file)
    except FileNotFoundError:
        config_data = {"username": "", "api_token": "", "output_file": ""}

    username_label = tb.Label(config_window, text=" Jira Username:", font=('Helvetica', 12), bootstyle='default')
    username_label.pack(pady=5)

    ToolTip(username_label, text="Enter your Jira username")

    username_entry = tb.Entry(config_window, width=30, font=('Helvetica', 12), bootstyle='info')
    username_entry.pack(pady=5)
    username_entry.insert(0, config_data["username"])

    token_label = tb.Label(config_window, text="Jira API Token:", font=('Helvetica', 12), bootstyle='default')
    token_label.pack(pady=5)

    ToolTip(token_label, text="Enter your Jira API token")

    token_entry = tb.Entry(config_window, show="*", width=30, font=('Helvetica', 12), bootstyle='info')
    token_entry.pack(pady=5)
    token_entry.insert(0, config_data["api_token"])

    # Output File
    output_label = tb.Label(config_window, text="Output Path:", font=('Helvetica', 12), bootstyle='default')
    output_label.pack(pady=5)

    ToolTip(output_label, text="Select the folder where the output file will be saved")

    output_entry = tb.Entry(config_window, width=30, font=('Helvetica', 12), bootstyle='info')
    output_entry.pack(pady=5)
    output_entry.insert(0, config_data["output_file"])

    # Select Folder Button
    select_folder_button = tb.Button(config_window, text="Select Folder", bootstyle='info', command=select_output_folder)
    select_folder_button.pack(pady=5)

    def save_config():
        # Save configuration to JSON file
        new_config_data = {
            "username": username_entry.get(),
            "api_token": token_entry.get(),
            "output_file": output_entry.get()
        }

        with open(CONFIG_FILE, "w") as file:
            json.dump(new_config_data, file)

        messagebox.showinfo("Config Saved", "Configuration saved successfully.")
        config_window.destroy()

    save_button = tb.Button(config_window, text="Save", bootstyle='success-outline', command=save_config)
    save_button.pack(side=tb.BOTTOM, pady=15)


def select_output_folder():
    folder_selected = filedialog.askdirectory()
    output_entry.delete(0, tb.END)
    output_entry.insert(0, folder_selected)


def delete_file(file_path):
    try:
        os.remove(file_path)
        print(f"{file_path} has been deleted.")
    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")


icon_path = os.path.join(base_dir, 'icon.png')

# Main window
root = tb.Window(
    title= "Sprint Processor", 
    themename='journal', 
    iconphoto=icon_path, 
    resizable=(False, False), 
    size=(400, 350)
    )

sprint_label = tb.Label(root, text="Jira Sprint ID:", font=('Helvetica', 14), bootstyle='default')
sprint_label.pack(pady=10)

ToolTip(sprint_label, text="Enter the Jira Sprint ID")

sprint_entry = tb.Entry(root, font=('Helvetica', 12), bootstyle='dark')
sprint_entry.pack(pady=10)

process_button = tb.Button(root, text="Start", bootstyle='success-outline', command=process_sprint)
process_button.pack(pady=10)

result_label = tb.Label(root, text="", font=('Helvetica', 12))
result_label.pack(pady=10)

config_button = tb.Button(root, text="Config.", bootstyle='secondary-outline', command=open_config_window)
config_button.pack(side=tb.BOTTOM, pady=15)

# Run the GUI
root.mainloop()
