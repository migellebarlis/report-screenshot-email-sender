import argparse
import datetime
import time
import os
import io
import smtplib
import html
from email.message import EmailMessage
from email.utils import make_msgid, parseaddr

from loguru import logger
from dotenv import load_dotenv

from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.discovery import build

from aspose.cells import Workbook
from aspose.cells.drawing import ImageType
from aspose.cells.rendering import ImageOrPrintOptions, SheetRender

# Get the path of the current file
current_path = os.path.dirname(os.path.abspath(__file__))

# Load settings from .env in the project folder (see .env.example)
load_dotenv(os.path.join(current_path, ".env"))

def get_env(name, default=None):
    value = os.getenv(name) or default
    if not value:
        raise SystemExit(f"Missing required setting {name} in .env (see .env.example)")
    return value

# Get date today
def set_report_date(date):
    """Set the date used for the worksheet, the block to find, and the email."""
    global DATE, DATE_FORMAT_1, DATE_FORMAT_2, DATE_FORMAT_3
    DATE = date
    DATE_FORMAT_1 = f"{DATE.strftime('%B')} {DATE.day}, {DATE.year}"
    DATE_FORMAT_2 = DATE.strftime("%Y-%m-%d")
    DATE_FORMAT_3 = DATE.strftime("%Y%m")

set_report_date(datetime.datetime.now())

# The service account only needs to read files from the shared drive
SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly"
]

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465

SHARED_DRIVE_ID = get_env("SHARED_DRIVE_ID")

FILE_PATH = get_env("FILE_PATH")

TEMP_IMAGE_FILE_NAME = get_env("TEMP_IMAGE_FILE_NAME", "image.png")

MESSAGE_TO = get_env("MESSAGE_TO")
MESSAGE_CC = os.getenv("MESSAGE_CC", "")
MESSAGE_FROM = get_env("MESSAGE_FROM")

# Gmail app password for MESSAGE_FROM; Google shows it with spaces, which are not part of it
GMAIL_APP_PASSWORD = get_env("GMAIL_APP_PASSWORD").replace(" ", "")

SERVICE_ACCOUNT_FILE = get_env("SERVICE_ACCOUNT_FILE", "sa-key.json")

EMAIL_SUBJECT = get_env("EMAIL_SUBJECT", "Sample")
REPORT_NAME = get_env("REPORT_NAME", "sample report")
REPORT_LINK = os.getenv("REPORT_LINK", "")
REPORT_LINK_LABEL = get_env("REPORT_LINK_LABEL", "Report Link")
SENDER_NAME = get_env("SENDER_NAME", "Sample")

# Width in pixels the screenshot is shown at in the email body
EMAIL_IMAGE_WIDTH = int(get_env("EMAIL_IMAGE_WIDTH", "700"))

def get_google_creds():
    key_path = os.path.join(current_path, SERVICE_ACCOUNT_FILE)
    if not os.path.exists(key_path):
        raise FileNotFoundError(f"Service account key not found: {key_path}")
    return service_account.Credentials.from_service_account_file(key_path, scopes=SCOPES)

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
            if not parents["files"]:
                raise FileNotFoundError(f"Folder '{file_parts[index]}' not found in shared drive")
            parent_id = parents["files"][0]["id"]
        files = service.files().list(q = f"name = '{file_parts[-1]}' and parents = '{parent_id}' and trashed = false", driveId=SHARED_DRIVE_ID, includeItemsFromAllDrives=True, corpora="drive", supportsAllDrives=True).execute()
        if not files["files"]:
            raise FileNotFoundError(f"File '{file_parts[-1]}' not found in shared drive")
        file_id = files["files"][0]["id"]
        return file_id
    except Exception as e:
        logger.error(f"Failed to get file id from google drive: {e}")
        raise


def download_from_drive(file_id):
    """Downloads a file
    Args:
    file_id: ID of the file to download
    Returns : the file contents as bytes.
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
        raise

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
    if sheet is None:
        raise ValueError(f"Worksheet '{DATE_FORMAT_3}' not found in {file_path}")

    # Set all margins of the worksheet to zero
    sheet.page_setup.left_margin = 0.0
    sheet.page_setup.bottom_margin = 0.0
    sheet.page_setup.top_margin = 0.0
    sheet.page_setup.right_margin = 0.0

    # max_data_row is 0-based, so add 1 to get the last used row number
    last_row = sheet.cells.max_data_row + 1
    row_start = 2
    sheet_range_date = sheet.cells.get(f"D{row_start}").display_string_value
    while sheet_range_date != DATE_FORMAT_1:
        row_start = row_start + 10
        if row_start > last_row:
            raise ValueError(f"Date '{DATE_FORMAT_1}' not found in column D of worksheet '{DATE_FORMAT_3}'")
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
    """Create and send an email message through Gmail SMTP using an app password."""
    try:
        message = EmailMessage()
        message["To"] = MESSAGE_TO
        if MESSAGE_CC:
            message["Cc"] = MESSAGE_CC
        message["From"] = MESSAGE_FROM

        message["Subject"] = f"{EMAIL_SUBJECT} - {DATE_FORMAT_1}"

        with open(image_path, "rb") as file:
            img_bytes = file.read()

        # Email clients block data: URLs, so the body references the image by Content-ID
        image_cid = make_msgid(domain="report.local")

        report_link_html = ""
        if REPORT_LINK:
            link = html.escape(REPORT_LINK)
            report_link_html = f"""<div>{html.escape(REPORT_LINK_LABEL)}:
            <a href="{link}" target="_blank">{link}</a>
            </div>"""

        message_content = f"""
            Hi,
            <br>
            <br>
            Please find below today's {REPORT_NAME} as of {DATE_FORMAT_1}.
            <br>
            <img alt="{REPORT_NAME}" src="cid:{image_cid[1:-1]}" width="{EMAIL_IMAGE_WIDTH}" style="width: {EMAIL_IMAGE_WIDTH}px; max-width: 100%; height: auto;" />
            <br>
            {report_link_html}
            <div>
            <br>
            Best,
            <br>
            {SENDER_NAME}
            </div>
            """
        message.set_content(message_content, subtype="html")
        message.add_related(img_bytes, maintype="image", subtype="png", cid=image_cid)

        message.add_attachment(img_bytes, maintype="image", subtype="png", filename=f"{DATE_FORMAT_2}.png")

        # send_message delivers to everyone in the To and Cc headers
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as smtp:
            # Log in with the bare address, even if MESSAGE_FROM includes a display name
            smtp.login(parseaddr(MESSAGE_FROM)[1], GMAIL_APP_PASSWORD)
            smtp.send_message(message)
    except Exception as e:
        logger.error(f"Failed to send email: {e}")
        raise

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
                raise SystemExit(1)

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

def date(value):
    return datetime.datetime.strptime(value, "%Y-%m-%d")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Screenshot a day's report block and email it.")
    parser.add_argument(
        "--date",
        type=date,
        help="report date as YYYY-MM-DD (default: today)",
    )
    args = parser.parse_args()
    if args.date:
        set_report_date(args.date)
    logger.info(f"Report date: {DATE_FORMAT_1}")

    # Number of tries and time to wait in between retries
    number_of_tries = 5
    waiting_time = 60

    # Run the main function with retry mechanism
    run_with_retry(number_of_tries, waiting_time)

