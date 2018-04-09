#MORE INFO: https://developers.google.com/resources/api-libraries/documentation/sheets/v4/python/latest/
#MORE INFO: https://developers.google.com/sheets/quickstart/python

import httplib2
import os
import io
from apiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from apiclient.http import MediaFileUpload
from apiclient import discovery
from apiclient import errors
import oauth2client
from oauth2client import client
from oauth2client import tools


SCOPES = 'https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = '/json/client_secret.json'
APPLICATION_NAME = 'Drive API Python Quickstart'

# Standard Credential Procedure
# ORIGINAL CODE: quickstart.py from Google Developer info
def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'drive-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

#RETURNS: Google Drive Service instance
def create_drive_service():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)
    return service

#RETURNS: Google Sheets Service instance - USE FOR EVERYTHING
#MORE INFO: https://developers.google.com/resources/api-libraries/documentation/sheets/v4/python/latest/
#MORE INFO: https://developers.google.com/sheets/quickstart/python

def create_sheet_service():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)
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


