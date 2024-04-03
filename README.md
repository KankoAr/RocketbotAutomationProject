# AutomationProject

This is a simple Python bot with selenium that reads input from an Excel file, fills out web forms, and sends emails.

## Installation

Ensure you have the required packages installed by running:

`pip install -r requirements.txt`

## Chromedriver

If need to download Chromedriver. You can find the latest version [here](https://googlechromelabs.github.io/chrome-for-testing/).

# Libraries

The following libraries are used in this project:
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [selenium](https://www.selenium.dev/documentation/en/)
- [smtplib](https://docs.python.org/3/library/smtplib.html)
## Note

Before running the code, please make sure to perform the following replacements in the script:
- Replace the email address (line #126) with your own.
- Replace the third-party password for Gmail (line #143) with your own.
