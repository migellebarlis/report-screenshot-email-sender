# report-screenshot-email-sender

Downloads an Excel report from a Google Shared Drive, takes a screenshot of today's section of the sheet, and emails it through Gmail. The screenshot goes in the email body and is also attached.

## Quick start

```bash
# 1. Create the environment (Python 3.12, conda-forge only)
conda create -n report-screenshot-email-sender --override-channels -c conda-forge python=3.12 pip
conda activate report-screenshot-email-sender
pip install -r requirements.txt

# 2. Put credentials.json in the project root (see "Google setup" below)

# 3. Create .env and fill it in (see "Configuration" below)
cp .env.example .env

# 4. Run
python main.py
```

Or double-click **`run.command`** in Finder, or run `./run.command`. It finds the conda env's Python on its own, so you don't need to activate the env first. When the run ends, the window stays open so you can read the log.

The first time you double-click it, macOS may block it. To allow it, right-click → **Open**, or go to **System Settings → Privacy & Security → Open Anyway**.

You can also use `conda env create -f environment.yml`. It only works if you have accepted the Anaconda Terms of Service for the default channels (`conda tos accept ...`) or removed those channels from your conda config.

The first run opens a browser so you can sign in to Google. After you sign in, the script saves `token.json`, and later runs reuse it.

## Google setup (one time)

1. In the [Google Cloud Console](https://console.cloud.google.com/), create or pick a project.
2. Turn on the **Google Drive API** and the **Gmail API** (APIs & Services → Library).
3. Set up the **OAuth consent screen**. If the user type is *External*, add the sending account under **Test users**.
4. Create credentials: **Create credentials → OAuth client ID → Application type: Desktop app**.
5. Download the JSON and save it as `credentials.json` in the project root.

The script asks for these scopes:
- `https://www.googleapis.com/auth/drive`
- `https://mail.google.com/`

If you change `SCOPES` in `main.py`, delete `token.json` and sign in again.

`credentials.json` and `token.json` are in `.gitignore`. Never commit them.

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
| `MESSAGE_FROM` | yes | The Google account you sign in with. |
| `EMAIL_SUBJECT` | no | The subject line becomes `<EMAIL_SUBJECT> - October 10, 2026`. |
| `REPORT_NAME` | no | Used in the body: "Please find below today's `<REPORT_NAME>` as of …". |
| `REPORT_LINK` | no | A link to add to the email body. Leave blank to leave it out. |
| `SENDER_NAME` | no | The name in the sign-off. |
| `TEMP_IMAGE_FILE_NAME` | no | File name for the screenshot the script writes locally (default `image.png`). |

If a required setting is missing, the script stops right away and names the setting. `.env` is in `.gitignore`, so it won't be committed.

## What the spreadsheet must look like

The script looks for the right section of the sheet like this:

- The file must be a real `.xlsx` on Drive, not a native Google Sheet.
- It must have a worksheet named after the current month as `YYYYMM` (for example `202610`).
- Each day's section is a block of **10 rows across columns A–D**. The date sits in **column D**, starting at `D2` and then every 10 rows (`D2`, `D12`, `D22`, …).
- The date cell has to display exactly like `October 10, 2026` (full month name, day without a leading zero, year).

If today's date isn't in the sheet, the script keeps searching down the sheet and never stops, so make sure today's row exists before it runs.

## Notes

- **Aspose.Cells license:** `aspose-cells-python` is a paid library. Without a license, the screenshots carry an evaluation watermark and some features are limited.
- **Retries:** If an error reaches `main()`, the script retries up to 5 times, waiting 60 seconds between tries. Both numbers are set in the `__main__` block.
- **Scheduling:** To send the report every day, run `python main.py` from cron or launchd using the conda env's Python. You can get that path with `conda run -n report-screenshot-email-sender which python`.
