from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

import os 
import json 
import time

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import pandas as pd


from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment

# SharePoint site URL
def sharepoint_client(username: str, password: str, sharepoint_site_url: str) -> ClientContext:
    """
    Creates and returns a SharePoint client context.
    """
    # Authenticate to SharePoint
    ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(username, password))

    # Load and verify connection
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()

    print(f"Authenticated successfully. Site Title: {web.properties['Title']}")
    return ctx

def download_file_from_sharepoint(client: ClientContext, sharepoint_file_url: str) -> str:
    """
    Downloads a file from SharePoint and returns the local file path.
    Handles both cases where subfolders exist or only the root folder is used.
    """
    # Extract the root folder, folder path, and file name
    path_parts = sharepoint_file_url.split('/')
    DOCUMENT_LIBRARY = path_parts[0]  # Root folder name (document library)
    FOLDER_PATH = '/'.join(path_parts[1:-1]) if len(path_parts) > 2 else ''  # Subfolders inside root, or empty if none
    file_name = path_parts[-1]  # File name

    # Construct the local folder path inside the Documents folder
    documents_folder = os.path.join(os.path.expanduser("~"), "Documents", FOLDER_PATH) if FOLDER_PATH else os.path.join(os.path.expanduser("~"), "Documents", DOCUMENT_LIBRARY)

    # Ensure the folder exists
    if not os.path.exists(documents_folder):
        os.makedirs(documents_folder)

    # Define the download path inside the folder
    download_path = os.path.join(os.getcwd(), file_name)

    # Download the file from SharePoint
    with open(download_path, "wb") as local_file:
        file = (
            client.web.get_file_by_server_relative_path(sharepoint_file_url)
            .download(local_file)
            .execute_query()
        )
    # Define the maximum wait time (60 seconds) and check interval (1 second)
    wait_time = 60  # 60 seconds
    elapsed_time = 0
    check_interval = 1  # Check every 1 second


    # While loop to check if the file exists at `file_path`
    while not os.path.exists(download_path) and elapsed_time < wait_time:
        time.sleep(check_interval)  # Wait 1 second
        elapsed_time += check_interval

    # After the loop, check if the file still doesn't exist and raise an error
    if not os.path.exists(download_path):
        raise FileNotFoundError(f"File not found at {download_path} after waiting for {wait_time} seconds.")

    print(f"[Ok] file has been downloaded into: {download_path}")
    return download_path

def upload_file_to_sharepoint(client: ClientContext, sharepoint_file_url: str, local_file_path: str, orchestrator_connection: OrchestratorConnection):
    """
    Uploads the specified local file back to SharePoint at the given URL.
    Uses the folder path directly to upload files.
    """
    # Extract the root folder, folder path, and file name
    path_parts = sharepoint_file_url.split('/')
    DOCUMENT_LIBRARY = path_parts[0]  # Root folder name (document library)
    FOLDER_PATH = path_parts[1]
    file_name = os.path.basename(local_file_path)  # File name

    # Construct the server-relative folder path (starting with the document library)
    if FOLDER_PATH:
        folder_path = f"{DOCUMENT_LIBRARY}/{FOLDER_PATH}"
    else:
        folder_path = f"{DOCUMENT_LIBRARY}"

    # Get the folder where the file should be uploaded
    target_folder = client.web.get_folder_by_server_relative_url(folder_path)
    client.load(target_folder)
    client.execute_query()

    # Upload the file to the correct folder in SharePoint
    with open(local_file_path, "rb") as file_content:
        uploaded_file = target_folder.upload_file(file_name, file_content).execute_query()


    orchestrator_connection.log_info(f"[Ok] file has been uploaded to: {uploaded_file.serverRelativeUrl} on SharePoint")


def create_empty_excel(file_path):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Overdragelser"
    
    # Insert column headers
    headers = ["Oprindelig aktivitetsbehandler", "Sagens sagsbehandler", "Ny aktivitetsbehandler"]
    worksheet.append(headers)
    
    # Insert an empty row to ensure the table is valid
    worksheet.append(["", "", ""])
    
    # Create a table
    table_ref = "A1:C2"
    table = Table(displayName="OverdragelserTable", ref=table_ref)
    table_style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                 showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    table.tableStyleInfo = table_style
    worksheet.add_table(table)
    
    # Auto-adjust column widths
    for col in worksheet.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        col_letter = col[0].column_letter
        worksheet.column_dimensions[col_letter].width = max_length + 2
    
    # Set alignment
    alignment = Alignment(vertical="center", wrap_text=True)
    for row in worksheet.iter_rows():
        for cell in row:
            cell.alignment = alignment
    
    # Save the workbook
    workbook.save(file_path)



orchestrator_connection = OrchestratorConnection("NovaOpgaveFlytDispatcher", os.getenv('OpenOrchestratorSQL'), os.getenv('OpenOrchestratorKey'), None)
RobotCredentials = orchestrator_connection.get_credential("Robot365User")
username = RobotCredentials.username
password = RobotCredentials.password
sharepoint_site_base = orchestrator_connection.get_constant("AarhusKommuneSharePoint").value


# SharePoint site URL
SHAREPOINT_SITE_URL = f"{sharepoint_site_base}/Teams/tea-teamsite10149"

client = sharepoint_client(username, password, SHAREPOINT_SITE_URL)

excel_path = download_file_from_sharepoint(client, "Delte dokumenter/Aktivitetsoverdragelse/Aktivitetsoverdragelse.xlsx")

planner_df = pd.read_excel(excel_path, sheet_name="Overdragelser")

os.remove(excel_path)

# Step 1: Prepare the data for the queue with trimmed and cleaned values
data = tuple(
    json.dumps({
        "OprindeligAktivitetsbehandler": str(row["Oprindelig aktivitetsbehandler"]).strip(),
        "SagensSagsbehandler": str(row["Sagens sagsbehandler"]).strip(),
        "NyAktivitetsbehandler": str(row["Ny aktivitetsbehandler"]).strip()
    }) for _, row in planner_df.iterrows()
)

references = tuple(
    f"{str(row['Oprindelig aktivitetsbehandler']).strip()} + "
    f"{str(row['Sagens sagsbehandler']).strip()} -> "
    f"{str(row['Ny aktivitetsbehandler']).strip()}"
    for _, row in planner_df.iterrows()
)

# Step 2: Call bulk_create_queue_elements
if data:
    orchestrator_connection.bulk_create_queue_elements("NovaOpgaveFlyt", references=references, data=data)
    create_empty_excel("Aktivitetsoverdragelse.xlsx")
    upload_file_to_sharepoint(client, "Delte dokumenter/Aktivitetsoverdragelse", "Aktivitetsoverdragelse.xlsx", orchestrator_connection)
    os.remove("Aktivitetsoverdragelse.xlsx")