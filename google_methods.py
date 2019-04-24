#MORE INFO: https://developers.google.com/resources/api-libraries/documentation/sheets/v4/python/latest/
#MORE INFO: https://developers.google.com/sheets/quickstart/python
#INSTALLATION: https://developers.google.com/api-client-library/python/start/installation

import os
import io
import pickle
import os.path
"""from apiclient.http import MediaIoBaseUpload, MediaIoBaseDownload #CONFIRM?
from apiclient.http import MediaFileUpload #CONFIRM?
from apiclient import discovery #CONFIRM?
from apiclient import errors #CONFIRM?
"""
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

"""
import oauth2client
from oauth2client import client
from oauth2client import tools
from oauth2client import file
"""

SCOPES = 'https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'json/client_secret.json'
APPLICATION_NAME = 'Drive API Python Quickstart'

# Standard Credential Procedure
# ORIGINAL CODE: quickstart.py from Google Developer info
def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    obtains the new credentials.

    Returns:
        creds, the obtained credential.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('auth/token.pickle'):
        with open('auth/token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'auth/credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('auth/token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

#RETURNS: Google Drive Service instance
def create_drive_service():
    creds = get_credentials()
    service = build('drive', 'v3', credentials=creds)
    return service

#RETURNS: Google Sheets Service instance - USE FOR EVERYTHING
#MORE INFO: https://developers.google.com/resources/api-libraries/documentation/sheets/v4/python/latest/
#MORE INFO: https://developers.google.com/sheets/quickstart/python

def create_sheet_service():
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    return service

#Return range from sheet
#ARGS: range in A1 notation. OPTIONAL: spreadsheet id, sheet title (defaults: general accounting, invoice sheet)
#RETURNS: range object (dictionary with values noted in list stored in 'values' key).

general_accounting_id = '1GhovgewjuCoqJRdN-PrJvXv9Tq9INeKqgBH_s__xi2M'
def get_range(base_range, spreadsheet_id = general_accounting_id, sheet_title = 'Factures 2016'):
    service = create_sheet_service()
    range_request = "'" + sheet_title + "'!" + base_range
    return(service.spreadsheets().values().get(spreadsheetId = spreadsheet_id, range = range_request).execute())

#Set range in sheet
#ARGS: range in A1 notation, values (list of lists of values - [[row values],[column values]])
#OPTIONAL ARGS: spreadsheet id, sheet title (defaults: general accounting, invoice sheet)
#RETURNS: None

def set_range(base_range, values, spreadsheet_id = general_accounting_id, sheet_title = 'Factures 2016'):
    service = create_sheet_service()
    range_request = "'" + sheet_title + "'!" + base_range
    request_body = {}
    request_body['range'] = range_request
    request_body['values'] = values
    service.spreadsheets().values().update(spreadsheetId = spreadsheet_id, range = range_request,
                                           body = request_body, valueInputOption = 'USER_ENTERED').execute()

#create new file from template
#ORIGINAL CODE: https://developers.google.com/drive/v2/reference/files/copy

def copy_file(origin_file_id):
    """Copy an existing file.
    
    Args:
    origin_file_id: ID of the origin file to copy.

    Returns:
    The copied file if successful, None otherwise.
  """
    service = create_drive_service()
    copied_file = {'title': "default"}
    try:
        return service.files().copy(
            fileId=origin_file_id, body=copied_file).execute()
    except errors.HttpError:
        print ('An error occurred: %s' % error)
        return None

#Update Google Drive file metadata - NOTE: NOT CURRENTLY NEEDED
#ORIGINAL CODE: https://developers.google.com/drive/v2/reference/files/update

def update_file(file_id, new_name = "", new_title = "", new_description = "", new_mime_type = "",
                new_filename = "", new_revision = "", service = "default"):
    """Update an existing file's metadata and content.

    Args:
    service: Drive API service instance.
    file_id: ID of the file to update.
    new_title: New title for the file.
    new_description: New description for the file.
    new_mime_type: New MIME type for the file.
    new_filename: Filename of the new content to upload.
    new_revision: Whether or not to create a new revision for this file.
    Returns:
    Updated file metadata if successful, None otherwise.
    """
    service = create_drive_service()
    try:
        # First retrieve the file from the API.
        file = service.files().get(fileId=file_id).execute()

        # File's new metadata.
        if new_name:
            file['name'] = new_name
        if new_title:
            file['title'] = new_title
        if new_description:
            file['description'] = new_description
        if new_mime_type:
            file['mimeType'] = new_mime_type
        print(file)

        """        
        # File's new content.
        media_body = MediaFileUpload(
            new_filename, mimetype=new_mime_type, resumable=True)

        # Send the request to the API.
        updated_file = service.files().update(
            fileId=file_id,
            body=file,
            newRevision=new_revision,
            media_body=media_body
            ).execute()
        return updated_file
        """        
    except errors.HttpError:
        print ('An error occurred: %s' % error)
        return None

#Get ID from File Name
def get_file_id(file_name):
    """
    ARGS: File name from Google Drive
    RETURNS: File ID
    """
    service = create_drive_service()
    try:
        # First retrieve the file from the API.
        file = service.files().list(q="name = '" + file_name + "'").execute()

        # Return file ID.
        try:
            return(file['files'][0]['id'])
        except IndexError:
            return None
    
    except errors.HttpError:
        print ('An error occurred: %s' % error)
        return None

#Download file from Google Drive as PDF

def download_pdf(file_id = '1REZ_1wDyAhh7W6EFqPR65XVeeeB0XSEs0s7DNpiwe0g', new_title = 'default.pdf'):
    """
    Downloads a file from Google Drive in PDF format.
    Note: Possible to change file type, such as xlsx.
    See: https://developers.google.com/drive/v3/web/manage-downloads

    Arguments: file_id(str)
    Action: Create file in PDF format on local computer.
    Returns: Nothing
    """
    print("Debugging download_PDF")
    service = create_drive_service()
    request = service.files().export_media(fileId=file_id,
                                                 mimeType='application/pdf')
 
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    count = 0
    while done is False:
        status, done = downloader.next_chunk()
        count += 1
        print ("download complete")
        if count > 10:
            return None
    with io.open(new_title + '.pdf', 'wb') as file:
        fh.seek(0)
        file.write(fh.read())

#Download Spreadsheet Info

def spreadsheet_info(spreadsheet_Id):
    service = create_sheet_service()
    file = service.spreadsheets().get(spreadsheetId=spreadsheet_Id).execute()
    return file

# Update Spreadsheet Info
# ORIGINAL CODE: https://developers.google.com/sheets/guides/batchupdate
# More examples available at above URL
# Takes two arguments:
# 1. Spreadsheet ID
# 2. List of dicts in following format:
"""
"updateSpreadsheetProperties":
    {
    "properties" : {"title" : "My New Title"},
    "fields" : "title"
    }
"""
# Final Basic Request format for reference (output generated by this function):
"""
{
  "requests": [{
      "updateSpreadsheetProperties": {
          "properties": {"title": "My New Title"},
          "fields": "title"
        }
    }]
}
"""

def update_spreadsheet(spreadsheet_id,request_array):
    service = create_sheet_service()
    spreadsheet_id = spreadsheet_id #WTF - prob useless
    requests = []

    for request in request_array:
        requests.append(request)

    batchUpdateRequest = {'requests': requests}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                       body=batchUpdateRequest).execute()


