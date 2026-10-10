# report-screenshot-email-sender

Downloads an Excel report from a Google Shared Drive, takes a screenshot of today's section of the sheet, and emails it through Gmail. The screenshot goes in the email body and is also attached.

## Quick start

```bash
# 1. Create the environment (Python 3.12, conda-forge only)
conda create -n report-screenshot-email-sender --override-channels -c conda-forge python=3.12 pip
conda activate report-screenshot-email-sender
pip install -r requirements.txt

# 2. Put sa-key.json in the project root and create a Gmail app password (see "Google setup" below)

# 3. Create .env and fill it in (see "Configuration" below)
cp .env.example .env

# 4. Run
python main.py
```

Or double-click **`run.command`** in Finder, or run `./run.command`. It finds the conda env's Python on its own, so you don't need to activate the env first. When the run ends, the window stays open so you can read the log.

To send the report for another day, for example to test with a block that's already in the sheet, pass a date:

```bash
./run.command --date 2026-10-08      # or: python main.py --date 2026-10-08
```

When you double-click `run.command`, it asks for a date first. Press Enter to use today. The date selects the worksheet (`YYYYMM`), the block to screenshot, the email subject, and the attachment name.

The first time you double-click it, macOS may block it. To allow it, right-click → **Open**, or go to **System Settings → Privacy & Security → Open Anyway**.

You can also use `conda env create -f environment.yml`. It only works if you have accepted the Anaconda Terms of Service for the default channels (`conda tos accept ...`) or removed those channels from your conda config.

The script never opens a browser or asks you to sign in, so it can run unattended on a schedule.

## Google setup (one time)

The script signs in to Google two ways: a **service account** to read the file from Drive, and a **Gmail app password** to send the email.

### Drive: service account

1. In the [Google Cloud Console](https://console.cloud.google.com/), create or pick a project.
2. Turn on the **Google Drive API** (APIs & Services → Library).
3. Go to **IAM & Admin → Service Accounts → Create service account**. It doesn't need any project roles.
4. Open the service account, then go to **Keys → Add key → Create new key → JSON**. Save the downloaded file as `sa-key.json` in the project root.
5. Copy the service account's email address (it looks like `name@project.iam.gserviceaccount.com`). In Google Drive, open the Shared Drive, choose **Manage members**, and add that email as a **Viewer**.

If key creation is blocked, your Google Workspace admin has turned off service account keys (the `iam.disableServiceAccountKeyCreation` policy). Ask them to allow it for this project.

The service account only asks for read-only Drive access (`https://www.googleapis.com/auth/drive.readonly`).

### Email: Gmail app password

1. Sign in as the sending account (`MESSAGE_FROM`) and turn on **2-Step Verification**. App passwords require it.
2. Go to [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords), create an app password, and copy the 16-character code.
3. Put it in `.env` as `GMAIL_APP_PASSWORD`. Spaces are fine; the script removes them.

If the app passwords page says the setting isn't available, either 2-Step Verification is off, or your Workspace admin has disabled app passwords for the account.

`sa-key.json` and `.env` are in `.gitignore`. Never commit them. Anyone with either one can read the Drive or send email as the account.

## Configuration

All settings are in a `.env` file in the project root, which `main.py` reads when it starts. To create it, copy the template and fill it in:

```bash
cp .env.example .env
```

| Setting | Required | What to set |
| --- | --- | --- |
| `SHARED_DRIVE_ID` | yes | ID of the Shared Drive. It's the last part of the drive's URL: `drive.google.com/drive/folders/<ID>`. |
| `FILE_PATH` | yes | Path to the `.xlsx` file inside the Shared Drive, with folders separated by `/` (for example `Reports/2026/daily.xlsx`). |
| `MESSAGE_TO` | yes | Recipients. Separate multiple addresses with commas. |
| `MESSAGE_CC` | no | CC recipients. Leave blank to send without CC. |
| `MESSAGE_FROM` | yes | The Gmail or Google Workspace address the report is sent from. |
| `GMAIL_APP_PASSWORD` | yes | App password for `MESSAGE_FROM` (see "Email: Gmail app password"). |
| `EMAIL_SUBJECT` | no | The subject line becomes `<EMAIL_SUBJECT> - October 10, 2026`. |
| `REPORT_NAME` | no | Used in the body: "Please find below today's `<REPORT_NAME>` as of …". |
| `REPORT_LINK` | no | A link to add to the email body. Leave blank to leave it out. |
| `REPORT_LINK_LABEL` | no | Label shown before the link, as `<label>: <link>` (default `Report Link`). |
| `SENDER_NAME` | no | The name in the sign-off. |
| `SERVICE_ACCOUNT_FILE` | no | Service account key file, relative to the project folder (default `sa-key.json`). |
| `EMAIL_IMAGE_WIDTH` | no | Width in pixels the screenshot is shown at in the email body (default `700`). The attachment stays full size. |
| `TEMP_IMAGE_FILE_NAME` | no | File name for the screenshot the script writes locally (default `image.png`). |

If a required setting is missing, the script stops right away and names the setting. `.env` is in `.gitignore`, so it won't be committed.

## What the spreadsheet must look like

The script looks for the right section of the sheet like this:

- The file must be a real `.xlsx` on Drive, not a native Google Sheet.
- It must have a worksheet named after the current month as `YYYYMM` (for example `202610`).
- Each day's section is a block of **10 rows across columns A–D**. The date sits in **column D**, starting at `D2` and then every 10 rows (`D2`, `D12`, `D22`, …).
- The date cell has to display exactly like `October 10, 2026` (full month name, day without a leading zero, year).

If today's date or this month's worksheet isn't in the file, the run fails with an error that says what's missing.

## Notes

- **Aspose.Cells license:** `aspose-cells-python` is a paid library. Without a license, the screenshots carry an evaluation watermark and some features are limited.
- **Retries:** If any step fails, the script retries up to 5 times, waiting 60 seconds between tries. Both numbers are set in the `__main__` block. If every try fails, it exits with code 1.
- **Scheduling:** To send the report every day, run `python main.py` from cron or launchd using the conda env's Python. You can get that path with `conda run -n report-screenshot-email-sender which python`.
