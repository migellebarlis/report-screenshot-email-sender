import datetime
import time
import os
import io
import base64
from email.message import EmailMessage

from PIL import Image
from loguru import logger

import google.auth
from google.oauth2 import service_account
from googleapiclient.http import MediaFileUpload
from googleapiclient.http import MediaIoBaseDownload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from aspose.cells import Workbook
from aspose.cells.drawing import ImageType
from aspose.cells.rendering import ImageOrPrintOptions, SheetRender

# Get the path of the current file
current_path = os.path.dirname(os.path.abspath(__file__))

# Get date today
DATE = datetime.datetime.now()
DATE_FORMAT_1 = f"{DATE.strftime('%B')} {DATE.day}, {DATE.year}"
DATE_FORMAT_2 = DATE.strftime("%Y-%m-%d")
DATE_FORMAT_3 = DATE.strftime("%Y%m")

# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://mail.google.com/"
]

SHARED_DRIVE_ID = "sample_drive_id"

FILE_PATH = "sample.xlsx"

TEMP_IMAGE_FILE_NAME = "image.png"

MESSAGE_TO = "sample@gmail.com"
MESSAGE_CC = "sample@gmail.com"
MESSAGE_FROM = "sample@gmail.com"

def get_google_creds():
    creds = None
    # use this if service account
    # creds = service_account.Credentials.from_service_account_file(
    #     filename="sa-key.json",
    #     scopes=SCOPES)
    # use this if desktop app
    if os.path.exists(os.path.join(current_path, "token.json")):
        creds = Credentials.from_authorized_user_file(os.path.join(current_path, "token.json"), SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                os.path.join(current_path, "credentials.json"), SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(os.path.join(current_path, "token.json"), "w") as token:
            token.write(creds.to_json())
    return creds

def get_file_id_from_drive(file_path):
    creds = get_google_creds()

    try:
        file_id = None
        file_parts = file_path.split("/")
        if len(file_parts) == 0:
            return file_id

        # create drive api client
        service = build("drive", "v3", credentials=creds)
        parent_id = SHARED_DRIVE_ID
        for index in range(len(file_parts)-1):
            parents = service.files().list(q = f"name = '{file_parts[index]}' and parents = '{parent_id}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false", driveId=SHARED_DRIVE_ID, includeItemsFromAllDrives=True, corpora="drive", supportsAllDrives=True).execute()
            parent_id = parents["files"][0]["id"]
        files = service.files().list(q = f"name = '{file_parts[-1]}' and parents = '{parent_id}' and trashed = false", driveId=SHARED_DRIVE_ID, includeItemsFromAllDrives=True, corpora="drive", supportsAllDrives=True).execute()
        file_id = files["files"][0]["id"]
        return file_id
    except Exception as e:
        logger.error(f"Failed to get file id from google drive: {e}")


def download_from_drive(file_id):
    """Downloads a file
    Args:
    file_id: ID of the file to download
    Returns : IO object with location.

    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    creds = get_google_creds()

    try:
        # create drive api client
        service = build("drive", "v3", credentials=creds)

        # pylint: disable=maybe-no-member
        # request = service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        request = service.files().get_media(fileId=file_id)
        file = io.BytesIO()
        downloader = MediaIoBaseDownload(file, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            logger.info(f"Downloading progress: {int(status.progress() * 100)}")
        return file.getvalue()

    except Exception as e:
        logger.error(f"Failed to download file from google drive: {e}")

def extract_img_from_worksheet(file_path):
    file_id = get_file_id_from_drive(file_path)
    file = download_from_drive(file_id)

    # Create an instance of the Workbook class.
    workbook = Workbook(io.BytesIO(file))

    # Set OnePagePerSheet option as true
    options = ImageOrPrintOptions()
    options.one_page_per_sheet = True
    options.image_type = ImageType.PNG

    # Get the first worksheet
    sheet = workbook.worksheets.get(DATE_FORMAT_3)

    # Set all margins of the worksheet to zero
    sheet.page_setup.left_margin = 0.0
    sheet.page_setup.bottom_margin = 0.0
    sheet.page_setup.top_margin = 0.0
    sheet.page_setup.right_margin = 0.0

    row_start = 2
    sheet_range_date = sheet.cells.get(f"D{row_start}").display_string_value
    while sheet_range_date != DATE_FORMAT_1:
        row_start = row_start + 10
        sheet_range_date = sheet.cells.get(f"D{row_start}").display_string_value
    row_start = row_start - 1
    row_end = row_start + 9
    sheet.page_setup.print_area = f"A{row_start}:D{row_end}"

    # Take the image of your worksheet
    sr = SheetRender(sheet, options)
    image_path = os.path.join(current_path, TEMP_IMAGE_FILE_NAME)
    sr.to_image(0, image_path)

    return image_path

def send_email(image_path):
    """Create and send an email message
    Print the returned  message id
    Returns: Message object, including message id

    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """
    creds = get_google_creds()

    try:
        service = build("gmail", "v1", credentials=creds)
        message = EmailMessage()
        message["To"] = MESSAGE_TO
        message["Cc"] = MESSAGE_CC
        message["From"] = MESSAGE_FROM

        message["Subject"] = f"Sample - {DATE_FORMAT_1}"

        with Image.open(image_path) as img:
            # Save the image data to a bytes buffer in the correct format
            buffer = io.BytesIO()
            img.save(buffer, format="png")
            img_bytes = buffer.getvalue()
            
            # Encode the bytes to a Base64 string
            base64_encoded_bytes = base64.b64encode(img_bytes)
            base64_string = base64_encoded_bytes.decode('utf-8')
            
            # Construct the full Data URL
            data_url = f"data:image/png;base64,{base64_string}"

        message_content = f"""
            Hi,
            <br>
            <br>
            Please find below today's sample report as of {DATE_FORMAT_1}.
            <br>
            <img alt="image.png" src="{data_url}" />
            <br>
            <div>Sample Report Link:
            <a href="https://static.wikia.nocookie.net/crayonshinchan/images/6/6d/SHIN.png/revision/latest?cb=20200609184539" target="_blank">Click Me!</a>
            <br>
            <br>
            Best,
            <br>
            Sample
            </div>
            """
        message.set_content(message_content, subtype="html")

        with open(image_path, "rb") as file:
            attachment_data = file.read()
        message.add_attachment(attachment_data, maintype="imageg", subtype="png", filename=f"{DATE_FORMAT_2}.png")

        # encoded message
        encoded_message = base64.urlsafe_b64encode(bytes(message.as_string(), "utf-8")).decode("utf-8")

        create_message = {"raw": encoded_message}
        # pylint: disable=E1101
        send_message = (
            service.users()
            .messages()
            .send(userId="me", body=create_message)
            .execute()
        )
    except Exception as e:
        logger.error(f"Failed to send email: {e}")

def run_with_retry(num_tries: int, wait_time: float):
    """
    A helper function that runs the main function with retry mechanism.

    :param num_tries: The number of tries to run the main function.
    :param wait_time: The amount of time to wait between retries (in seconds).
    """
    for i in range(num_tries):
        try:
            main()
            break
        except Exception as e:
            logger.error(f"An error occurred: {e}")
            if i < num_tries - 1:
                logger.info(f"Waiting {wait_time} seconds before retrying...")
                time.sleep(wait_time)
            else:
                logger.error(f"Failed to run the program after {num_tries} tries.")

def main():
    """
    This main function exports worksheet range to image and sends an email.
    """
    # Starting the program
    logger.info("Starting the program")

    try:
        logger.info(f"Exporting worksheet range to image")
        image_path = extract_img_from_worksheet(FILE_PATH)
        logger.success(f"Worksheet range export to image has been successful")

        logger.info(f"Sending email")
        send_email(image_path)
        logger.success(f"Email has been successfully sent")

    except Exception as e:
        logger.error(f"An error occurred: {e}")
        raise

if __name__ == "__main__":
    # Number of tries and time to wait in between retries
    number_of_tries = 5
    waiting_time = 60

    # Run the main function with retry mechanism
    run_with_retry(number_of_tries, waiting_time)

