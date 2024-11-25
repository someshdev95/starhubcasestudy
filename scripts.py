# Import libraries 
import pandas as pd
import os
import win32com.client
import nzpy


# Function to download attachments
def download_attachments(mail_item):
    """
    Download all attachments from the given email (mail_item) in Outlook.

    Parameters:
    -----------
    mail_item : object
        The Outlook mail item from which attachments are to be downloaded.
        This should be an instance of an Outlook email object that has attachments.

    Returns:
    --------
    None
        The function saves the attachments to the specified directory.

    Side Effects:
    -------------
    - Creates a directory (`/attachments`) if it does not already exist.
    - Saves each attachment in the specified directory.
    """
    for attachment in mail_item.Attachments:
        # Define the file path where the attachment will be saved
        # Directory to save attachments
        attachment_directory = "/attachments"  # Modify with your desired directory
        if not os.path.exists(attachment_directory):
            os.makedirs(attachment_directory)

        file_path = os.path.join(attachment_directory, attachment.FileName)

        # Save the attachment
        attachment.SaveAsFile(file_path)
        print(f"Attachment {attachment.FileName} saved to {file_path}")


# Function to write to database
def write_to_netezza(df, table_name, conn):
    """
    Inserts a pandas DataFrame into the specified Netezza database table.

    Parameters:
    -----------
    df : pandas.DataFrame
        The DataFrame containing the data to be inserted into the Netezza table.

    table_name : str
        The name of the target table in the Netezza database.

    conn : object
        The active connection object to the Netezza database.
        It is typically created using libraries like `pyodbc` or `nzpy`.

    Returns:
    --------
    None
        The function does not return anything; it inserts the data into the database.
    """
    cursor = conn.cursor()

    # Create the insert query dynamically based on the DataFrame columns
    columns = ', '.join(df.columns)
    values = ', '.join(['%s'] * len(df.columns))

    insert_query = f"INSERT INTO {table_name} ({columns}) VALUES ({values})"

    # Iterate over the DataFrame rows and insert them
    for row in df.itertuples(index=False, name=None):
        cursor.execute(insert_query, row)

    conn.commit()
    cursor.close()


# We define functions for each table, making it more streamlined to add on more steps in the future
def clean_active(active, customer, date=None):
    """
    Filters and processes active service records as of the previous day.

    Parameters:
    -----------
    active : pandas.DataFrame
        DataFrame with service records as of the previous day.

    customer : pandas.DataFrame
        DataFrame with customer details.

    date : str or datetime-like, optional
        The reference date for filtering. Defaults to the current date.

    Returns:
    --------
    pandas.DataFrame
        A filtered and cleaned DataFrame with active services, merged with customer details.
    """
    if date is None:
        date = pd.to_datetime('today')
    else:
        date = pd.to_datetime(date)

    # Rename report date for better readability
    active = active.rename(columns={'REPORT_DATE': 'SNAPSHOT_DATE'})

    # Change all ID columns to string format for uniformity
    active[['CUSTOMER_ID', 'SERVICE_ID', 'SERVICE_NAME']] = active[['CUSTOMER_ID', 'SERVICE_ID', 'SERVICE_NAME']].astype(str)

    # Convert SNAPSHOT_DATE to datetime format and then to string
    active['SNAPSHOT_DATE'] = pd.to_datetime(active['SNAPSHOT_DATE'], format='%d/%m/%Y')
    active['SNAPSHOT_DATE'] = active['SNAPSHOT_DATE'].dt.strftime('%Y-%m-%d')

    # Filter to only include dates that are less than or equal to yesterday's date
    date_gap = pd.to_datetime('today') - pd.DateOffset(days=1)
    active = active[active['SNAPSHOT_DATE'] <= date_gap.strftime('%Y-%m-%d')]

    # Filter for active status and drop the column after
    active = active[active['SUBSCRIPTION_STATUS'] == 'active']
    active = active.drop('SUBSCRIPTION_STATUS', axis=1)

    # Merge with the customer table to get customer info
    customer[['CUSTOMER_ID']] = customer[['CUSTOMER_ID']].astype(str)

    # Remove any duplicate CUSTOMER_ID from the customer table since it's by transaction date
    customer = customer.drop_duplicates(subset=['CUSTOMER_ID'])

    # Remove date column
    customer = customer.drop('REPORT_DATE', axis=1)

    # Perform merge
    final_active = pd.merge(active, customer, on=['CUSTOMER_ID'], how='left')

    return final_active


def clean_order(order, service):
    """
    Cleans and merges order and service DataFrames based on service ID and report date.

    Parameters:
    -----------
    order : pandas.DataFrame
        DataFrame containing order data.

    service : pandas.DataFrame
        DataFrame containing service data.

    Returns:
    --------
    pandas.DataFrame
        A DataFrame resulting from an inner join of `order` and `service` on 'SERVICE_ID' and 'REPORT_DATE'.
    """
    # Convert REPORT_DATE to proper datetime format
    order['REPORT_DATE'] = pd.to_datetime(order['REPORT_DATE'])
    order['REPORT_DATE'] = order['REPORT_DATE'].dt.strftime('%Y-%m-%d')

    service['REPORT_DATE'] = pd.to_datetime(service['REPORT_DATE'])
    service['REPORT_DATE'] = service['REPORT_DATE'].dt.strftime('%Y-%m-%d')

    # Change ID columns to string format
    service[['SERVICE_ID']] = service[['SERVICE_ID']].astype(str)
    order[['SERVICE_ID']] = order[['SERVICE_ID']].astype(str)

    # Perform inner join on SERVICE_ID and REPORT_DATE
    df_clean = order.merge(service, on=['SERVICE_ID', 'REPORT_DATE'], how='left')

    return df_clean


# We run the pipeline

# Set up your Outlook Application
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Access the inbox folder
folder_name = "Inbox"  # You can change this to any folder, e.g., "Sent", or custom folder names
folder = namespace.GetDefaultFolder(6).Folders.Item(folder_name)  # 6 is for Inbox

# Get all the mails in the folder
messages = folder.Items
messages.Sort("[ReceivedTime]", True)  # Sort by received time, latest first

# Loop through emails to download attachments
for message in messages:
    try:
        # Check if the email has attachments
        if message.Attachments.Count > 0:
            print(f"Processing email: {message.Subject}")
            download_attachments(message)  # Download attachments
    except Exception as e:
        print(f"Error processing email: {str(e)}")

# Start preprocessing the files, assumption here is that the files are within the same directory
service = pd.read_csv('Raw Service.csv')
customer = pd.read_csv('Raw Customer.csv')
order = pd.read_csv('Raw Orders.csv')
active = pd.read_csv('Raw Active.csv')

active_final = clean_active(active, customer)
order_final = clean_order(order, service)

# Netezza connection details to write to DB
conn = nzpy.connect(
    host='host',
    user='username',
    password='password',
    database='database',
    port=5480
)

write_to_netezza(active_final, 'active_final', conn)
write_to_netezza(order_final, 'order_final', conn)
