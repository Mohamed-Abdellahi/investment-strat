# %%
"""
    Author: Mohamed-Abdellahi MOHAMED-ABDELLAHI
    mohamed-abdellahi.mohamed-abdellahi@dauphine.eu
    Date: 2024 Q4
    Role: Credit & algo trading Intern
"""


# check openpyxl
import os
import pandas as pd

import win32com.client
import win32com.client as win32
import os
from datetime import datetime, timedelta
from colorama import Fore, Style

# %%

def import_emails(folder_name, subject_keyword, save_folder, start_date, end_date):
    """
    

    This function connects to Outlook, filters emails by date and subject keyword,
    displays the email information, and saves the attachments to the specified folder.

    Parameters:
    folder_name (str): The name of the subfolder to search within.
    subject_keyword (str): The keyword to search for in the email subjects.
    start_date (str): The start date for filtering emails in the format 'dd-mm-yyyy'.
    end_date (str): The end date for filtering emails in the format 'dd-mm-yyyy'.
    save_folder (str): The path to the folder where attachments will be saved. eg: r"P:\rutger\REPO\historical"

    
    """
    # Connection
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.Folders.Item("mohamed.abdellahi@ca-cib.com")

    inbox = inbox.Folders["Boîte de réception"]

    # a good practise: Check if the subfolder exists
    try:
        folder = inbox.Folders[folder_name]  
    except Exception as e:
        print(f"{Fore.RED}Error: The folder '{folder_name}' was not found.")
        print(f"Error details: {e}{Style.RESET_ALL}")
        return

    # Convert dates to datetime objects
    try:
        start_date = datetime.strptime(start_date, "%d-%m-%Y")
        end_date = datetime.strptime(end_date, "%d-%m-%Y")
    except ValueError as e:
        print(f"{Fore.RED}Error: Invalid date format. Please use 'dd-mm-yyyy'.{Style.RESET_ALL}")
        print(f"Error details: {e}")
        return

    # Format dates for the filter
    start_date_str = start_date.strftime('%d/%m/%Y 00:00 AM')
    end_date_str = end_date.strftime('%d/%m/%Y 11:59 PM')

    # Display the dates used in the filter
    print(f"Filter date range: {start_date_str} to {end_date_str}")
    print(f"Filter date: [ReceivedTime] >= '{start_date_str}' AND [ReceivedTime] <= '{end_date_str}'")

    # Retrieve emails in the folder
    messages = folder.Items
    messages = messages.Restrict(f"[ReceivedTime] >= '{start_date_str}' AND [ReceivedTime] <= '{end_date_str}'")

    if messages.Count == 0:
        print("No items found with the specified date filter.")
        return

    print(f"Emails found between {start_date.strftime('%d-%m-%Y')} and {end_date.strftime('%d-%m-%Y')}:")
    print("-" * 50)

    # Iterate through the filtered emails
    for msg in messages:
        if subject_keyword.lower() in msg.Subject.lower():
            print(f"Subject: {msg.Subject}")
            print(f"Received Time: {msg.ReceivedTime}")
            print("Attachments:")
            for attachment in msg.Attachments:
                attachment_path = os.path.join(save_folder, attachment.FileName)
                try:
                    attachment.SaveAsFile(attachment_path)
                    print(f"  - {attachment.FileName} saved to {attachment_path}")
                except Exception as e:
                    print(f"{Fore.RED}Error saving attachment {attachment.FileName}: {e}{Style.RESET_ALL}")
            print("-" * 50)

#####------""""



##### sending the email
def send_email(subject, body, attachment_path, recipients, cc_recipients= None):
    
    outlook = win32.Dispatch('outlook.application')

    
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = body


    mail.To = "; ".join(recipients)

    if cc_recipients:
        mail.CC = "; ".join(cc_recipients)
    # Attach the file
    mail.Attachments.Add(attachment_path)

    
    #mail.Display(True)
    mail.Send()

# Function to read Excel files from the specified folder
def read_excel_files(folder_path):
    # The reading files takes about 80 sec for 20 files.. that’s a lot.. we should think of multiprocessing way or creating a database that chargees everything and then each day it only needs to update 
    all_data = []
    for filename in os.listdir(folder_path):
        if filename.startswith("RepoLevelReport 70833-F") and filename.endswith(".xlsx"):
            file_path = os.path.join(folder_path, filename)
            data = pd.read_excel(file_path, sheet_name="Apex for CDR")
            # Extract the pricing date from the filename
            pricing_date = datetime.strptime(filename.split()[-1].split('.')[0], '%Y-%m-%d').date()

            ## as in teh file receiveed, we don't have in the tabkle the date of repo pricing... 
            data['Pricing Date'] = pricing_date
            all_data.append(data)
    return pd.concat(all_data, ignore_index=True)

# Function to filter and clean the data
def filter_and_clean_data(data):
    # Filter rows where "Contra Account" is not empty and "Type" is "REP" or "REV"
    filtered_data = data[(data['Contra Account'].notna()) & (data['Type'].isin(['REP', 'REV']))]

    # important conversions 
    filtered_data['Pricing Date']= pd.to_datetime(filtered_data['Pricing Date'])
    
    filtered_data['Rate'] = pd.to_numeric(filtered_data['Rate'], errors= 'coerce')
    

    return filtered_data.reset_index(drop= True)

# Function to compute the variation of the rate
def compute_rate_variation(data, interval=1, start_date=None, end_date=None):
    #data['Pricing Date'] = pd.to_datetime(data['Pricing Date'])
    data.sort_values(by=['Trade No', 'Pricing Date'], inplace=True)

    if start_date and end_date:
        data = data[(data['Pricing Date'] >= start_date) & (data['Pricing Date'] <= end_date)]

    data['Previous Rate'] = data.groupby('Trade No')['Rate'].shift(periods=interval)

    # Shift the 'Pricing Date' column by the specified interval to get the previous pricing date
    data['Previous Pricing Date'] = data.groupby('Trade No')['Pricing Date'].shift(periods=interval)


    data['Rate Variation'] = data.groupby('Trade No')['Rate'].pct_change(periods=interval) * 100

    #data['Rate Variation'] = data['Rate Variation'].apply(lambda x: '' if pd.isna(x) else x)
    
    
    return data.reset_index(drop= True)

def get_variations(data, interval=1, start_date=None, end_date=None):
    
    data= compute_rate_variation(data, interval, start_date, end_date)

    most_recent_date= data['Pricing Date'].max()
    most_recent_data = data[data['Pricing Date'] == most_recent_date]

    return most_recent_data.reset_index(drop=True)


def get_cheapest_and_costliest(data, limit= 5, interval= 1):
    #data['Pricing Date']= pd.to_datetime(data['Pricing Date'])

    end_date= data['Pricing Date'].max()
    start_date= end_date - pd.Timedelta(days= interval-1)

    filtered_data= data[(data['Pricing Date'] >= start_date) & (data['Pricing Date'] <= end_date)]

    #filtered_data['Rate'] = pd.to_numeric(filtered_data['Rate'], errors= 'coerce')

    costiest= filtered_data.nlargest(limit, 'Rate')
    cheapest= filtered_data.nsmallest(limit, 'Rate')

    return costiest, cheapest


def identify_new_trades(data, interval=1):
    # Ensure 'Pricing Date' is in datetime format
    #data['Pricing Date'] = pd.to_datetime(data['Pricing Date'])

    # Sort the data by 'Trade No' and 'Pricing Date'
    data.sort_values(by=['Trade No', 'Pricing Date'], inplace=True)

    # Shift the 'Pricing Date' column by the specified interval to get the previous pricing date
    data['Previous Pricing Date'] = data.groupby('Trade No')['Pricing Date'].shift(periods=interval)

    # Identify new trades based on the specified interval
    data['New Trade'] = data['Previous Pricing Date'].isna()

    most_recent_date= data['Pricing Date'].max()
    most_recent_data = data[data['Pricing Date'] == most_recent_date]

    # Filter to return only the rows where it's a new trade
    new_trades = most_recent_data[most_recent_data['New Trade']]

    return new_trades


#  %%
##################################### THE MAIN #############################
###################                                       ##################
###########################################################################
# Main function to generate the report
def generate_repo_report(folder_path, output_folder):
    # Read and process data
    data = read_excel_files(folder_path)
    data = filter_and_clean_data(data)

    # Identify the most recent pricing date
    most_recent_date = data['Pricing Date'].max().strftime('%Y-%m-%d')

    # Generate the output filename
    output_file = os.path.join(output_folder, f'repo_reporting_{most_recent_date}.xlsx')

    # Ensure the output directory exists
    os.makedirs(output_folder, exist_ok=True)

    # Identify new trades
    new_trades = identify_new_trades(data)

    # Compute rate variations
    variation_data = compute_rate_variation(data)

    # Get the cheapest and most expensive repo rates
    cheapest, costliest = get_cheapest_and_costliest(variation_data)

    # Create a Pandas Excel writer using openpyxl as the engine
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write variation data to a separate sheet
        variation_data.to_excel(writer, sheet_name='Variation Data', index=False)

        # new trades to a separate sheet too
        new_trades.to_excel(writer, sheet_name='New Trades', index=False)

    
        # 5 cheapest and costliest rates 
        cheapest.to_excel(writer, sheet_name='Top 5 Cheapest Repo', index=False)

        
        costliest.to_excel(writer, sheet_name='Top 5 Most Expensive Repo', index=False)

        # Apply formatting to each sheet
        for sheet_name in writer.sheets:
            workbook = writer.book
            worksheet = workbook[sheet_name]
            worksheet.column_dimensions['A'].width = 20
            worksheet.column_dimensions['B'].width = 20
            worksheet.column_dimensions['C'].width = 20
            worksheet.column_dimensions['D'].width = 20
            worksheet.column_dimensions['E'].width = 20

    return variation_data, new_trades, cheapest, costliest
# %%

##################################### MAIN #######################################################

### importing the email

folder_name = "repo"  # Replace with the name of your subfolder 
subject_keyword = "70833 financing rates"

# files path
historical_folder = r"\\ldn.emea.cib\DFS01\FlowDesk\NON FLOW TRADING\repo\historical"
output_folder = r"\\ldn.emea.cib\DFS01\FlowDesk\NON FLOW TRADING\repo\daily_reporting"
## 

last10days= (datetime.today()-timedelta(days=10)).strftime('%d-%m-%Y')
today= datetime.today().strftime('%d-%m-%Y')
import_emails(folder_name, subject_keyword, historical_folder, last10days, today)


#variation_data, new_trades, cheapest, costliest = generate_repo_report(folder_path, output_path)


interval = 1  # For 1D variation
variation_data, new_trades, cheapest, costliest = generate_repo_report(historical_folder, output_folder)
print("Report generated successfully.")

# %%
## sending the email

subject = f"Daily covered bonds repo report {datetime.today().strftime('%Y-%m-%d')}"
body = (
    "Please find attached the daily repo report for covered bonds.\n"
    "If you have any questions or need further information, please do not hesitate to contact me.\n\n"
    "Best regards,\n"
    "Mohamed Abdellahi"
)

attachment_path = os.path.join(output_folder, f"repo_reporting_{datetime.today().strftime('%Y-%m-%d')}.xlsx")
report_date = datetime.today().strftime('%d-%m-%Y')  # This should be dynamically generated from your data
recipients = ["rutger.vorst@ca-cib.com"]  # List of recipients
cc_recipients= ["EUcredittradinginterns@ca-cib.com","mohamed.abdellahi@ca-cib.com"]

send_email(subject, body, attachment_path, recipients, cc_recipients)

# %%
