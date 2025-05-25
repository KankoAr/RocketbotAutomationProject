# AutomationProject: Excel to Web Form and Email Bot

This Python script automates the process of reading data from an Excel file, filling out a web form based on that data, and sending email notifications for specific conditions. It has been refactored for improved efficiency, robustness, and maintainability.

## Key Features and Improvements

*   **Efficient Web Interaction:** WebDriver (Chrome) is initialized only once if needed, rather than for each row, significantly speeding up processing for large files.
*   **Robust Form Filling:**
    *   Detailed error handling for each field during web form submission.
    *   Graceful handling of missing dropdown options, with clear logging of available options and attempted values (including normalization like lowercasing or stripping whitespace).
    *   Resilient parsing of submission confirmation messages.
*   **Secure Credential Management:** Email credentials are managed externally in a `config.ini` file, keeping them separate from the source code. An example file (`config_example.ini`) is provided.
*   **Comprehensive Diagnostic Logging:** Utilizes Python's `logging` module for all console output, providing:
    *   Timestamps for all actions.
    *   Severity levels (INFO, WARNING, ERROR).
    *   Detailed stack traces for errors, aiding in troubleshooting.
    *   Clear logging of data transformations (e.g., lowercasing for process name values, stripping whitespace for severity values) when matching dropdown options.
*   **Modular Design:** Code is organized into functions for clarity and reusability.

## Prerequisites

*   Python 3.x
*   Google Chrome browser installed.
*   Chromedriver (see section below).

## Installation

1.  **Clone the repository (if applicable) or download the script files.**
2.  **Install required Python packages:**
    Open a terminal or command prompt in the project directory and run:
    ```bash
    pip install -r requirements.txt
    ```
    This will install `openpyxl` (for Excel file reading) and `selenium` (for web automation).

## Chromedriver Setup

The script requires Chromedriver to interact with the Google Chrome browser.

1.  **Download Chromedriver:**
    *   Visit the official Chromedriver download page: [Chrome for Testing - WebDriver](https://googlechromelabs.github.io/chrome-for-testing/)
    *   Download the version of Chromedriver that matches **your installed Google Chrome browser version**.
2.  **Placement:**
    *   **Recommended:** Place the `chromedriver.exe` (or `chromedriver` on Linux/macOS) executable in a directory that is part of your system's PATH environment variable.
    *   **Alternative:** You can place the Chromedriver executable in the same directory as the `main.py` script.

## Configuration

### 1. Email Credentials (`config.ini`)

Email notifications are sent using credentials specified in a `config.ini` file.

1.  **Create `config.ini`:**
    *   In the project directory, find the file named `config_example.ini`.
    *   Make a copy of this file and rename the copy to `config.ini`.
2.  **Edit `config.ini`:**
    *   Open `config.ini` with a text editor.
    *   Under the `[SMTP]` section:
        *   Replace `your_email@gmail.com` with your actual Gmail address.
        *   Replace `your_app_password` with your Gmail **App Password**.
            *   **Important:** If you have 2-Step Verification enabled on your Gmail account (which is highly recommended), you **must** generate and use an App Password. Do not use your regular Gmail password.
            *   Instructions: [Google Account Help - Sign in with App Passwords](https://support.google.com/accounts/answer/185833)
    *   Example `config.ini` structure:
        ```ini
        [SMTP]
        email = your_actual_email@gmail.com
        password = your_generated_app_password
        ```
3.  **Security Warning:**
    *   The `config.ini` file contains sensitive credentials.
    *   **DO NOT commit `config.ini` to version control systems (e.g., Git).** If you use Git for this project, ensure `config.ini` is listed in your `.gitignore` file.

### 2. Excel File

*   The script expects an Excel file named `Base Seguimiento Observ Auditoría al_30042021.xlsx` in the same directory.
*   The structure of this Excel file (column order, sheet name) is implicitly defined in the `process_excel_file` function. Modify the script if your Excel file has a different name or structure.

## Running the Script

1.  Ensure all setup and configuration steps are complete.
2.  Open a terminal or command prompt in the project directory.
3.  Run the script:
    ```bash
    python main.py
    ```
4.  Observe the console output for logs from the script, indicating its progress and any errors.

## Libraries Used

*   **`openpyxl`**: For reading data from Excel files. (Installed via `requirements.txt`)
*   **`selenium`**: For browser automation and web form interaction. (Installed via `requirements.txt`)
*   **Standard Libraries (built-in with Python):**
    *   `smtplib`: For sending emails.
    *   `configparser`: For reading the `config.ini` file.
    *   `logging`: For structured and informative console output.
    *   `time`: For handling waits and delays (used implicitly by Selenium and for explicit waits).

```
