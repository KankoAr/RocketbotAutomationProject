import time
import smtplib
import logging 
import configparser
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def process_excel_file():
    excel_filepath = "Base Seguimiento Observ Auditoría al_30042021.xlsx" # This filename itself is in Spanish, but it's an external dependency.
    workbook = None
    driver = None

    try:
        workbook = load_workbook(excel_filepath)
        worksheet = workbook.active
        
        needs_webdriver = False
        for row_values in worksheet.iter_rows(min_row=2, values_only=True):
            if row_values[9] == 'Regularizado': # status in 10th column (index 9)
                needs_webdriver = True
                break
        
        if needs_webdriver:
            try:
                driver = webdriver.Chrome()
                logging.info("WebDriver initialized.")
            except Exception as e:
                logging.error(f"Error initializing WebDriver: {e}", exc_info=True)
                driver = None 
        
        for row_idx, current_row_cells in enumerate(worksheet.iter_rows(min_row=2), start=2):
            process_value_for_log = current_row_cells[0].value 
            try:
                process_name = current_row_cells[0].value
                observation = current_row_cells[1].value
                risk_type = current_row_cells[2].value
                severity = current_row_cells[3].value
                action_plan = current_row_cells[4].value 
                commitment_date = current_row_cells[5].value
                responsible_person = current_row_cells[6].value
                responsible_area = current_row_cells[7].value 
                responsible_person_email = current_row_cells[8].value
                status = current_row_cells[9].value
                
                if status == 'Regularizado': # This value 'Regularizado' is from the Excel data.
                    if driver:
                        logging.info(f"Processing row {row_idx} for web submission: {process_name}")
                        upload_information_to_form(driver, process_name, risk_type, severity, responsible_person, commitment_date, observation)
                    else:
                        logging.warning(f"Skipping web submission for row {row_idx} (process {process_name}) as WebDriver is not available.")
                elif status == 'Atrasado': # This value 'Atrasado' is from the Excel data.
                    logging.info(f"Processing row {row_idx} for email: {process_name}")
                    send_status_email(process_name, status, observation, commitment_date, responsible_person_email)
            except Exception as e:
                logging.error(f"Error processing row {row_idx} ({process_value_for_log if process_value_for_log else 'N/A'}): {e}", exc_info=True)
        
    except FileNotFoundError as e:
        logging.error(f"Error: Excel file '{excel_filepath}' not found.", exc_info=True)
    except Exception as e:
        logging.error(f"Error in process_excel_file function before row loop: {e}", exc_info=True)
    finally:
        if workbook:
            try:
                workbook.close()
                logging.info("Excel workbook closed.")
            except Exception as e:
                logging.error(f"Error closing workbook: {e}", exc_info=True)
        if driver:
            try:
                driver.quit()
                logging.info("WebDriver closed.")
            except Exception as e:
                logging.error(f"Error closing WebDriver: {e}", exc_info=True)

def upload_information_to_form(driver, original_process_name, risk_type, original_severity, responsible_person, commitment_date, observation):
    process_name_for_logging = original_process_name 
    try:
        driver.get("https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG")
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "process")))
        submission_successful = True

        # 1. Process Dropdown
        try:
            process_name_input_str = str(original_process_name) if original_process_name is not None else ""
            process_name_lower = process_name_input_str.lower()
            
            select_element = driver.find_element(By.ID, 'process')
            select_object = Select(select_element)
            process_found = False
            if original_process_name is not None: 
                for option in select_object.options:
                    if option.text.lower() == process_name_lower:
                        option.click()
                        process_found = True
                        break
            
            if not process_found:
                available_options = [opt.text for opt in select_object.options]
                if original_process_name is None:
                    logging.warning(f"Warning (Process '{process_name_for_logging}'): Input for 'process' was empty/None. No option selected. Available: {available_options}.")
                else:
                    log_msg = f"Warning (Process '{process_name_for_logging}'): Option '{process_name_lower}' (normalized from '{original_process_name}') not found in 'process' dropdown."
                    log_msg += f" Available: {available_options}. Attempting to continue."
                    logging.warning(log_msg)
        except (NoSuchElementException, TimeoutException) as e:
            logging.error(f"Critical error (Process '{process_name_for_logging}'): Could not find or interact with 'process' dropdown: {e}. Skipping this submission.", exc_info=True)
            return 

        # 2. Risk Type
        try:
            if risk_type is not None:
                 driver.find_element(By.ID, 'tipo_riesgo').send_keys(str(risk_type)) 
            else:
                logging.info(f"Information (Process '{process_name_for_logging}'): Field 'risk_type' is empty. Skipping.")
        except (NoSuchElementException, ElementNotInteractableException, TimeoutException) as e:
            logging.error(f"Error (Process '{process_name_for_logging}'): Could not interact with 'risk_type' field: {e}. Continuing to next field.", exc_info=True)
            submission_successful = False

        # 3. Severity Dropdown
        try:
            if original_severity is not None:
                severity_input_str = str(original_severity) 
                severity_stripped = severity_input_str.strip()
                severity_element = driver.find_element(By.ID, 'severidad') 
                severity_select_object = Select(severity_element)
                severity_found = False
                for option in severity_select_object.options:
                    if option.text == severity_stripped: 
                        option.click()
                        severity_found = True
                        break
                if not severity_found:
                    available_options = [opt.text for opt in severity_select_object.options]
                    log_msg = f"Warning (Process '{process_name_for_logging}'): Option '{severity_stripped}' (normalized from '{original_severity}') not found in 'severity' dropdown."
                    log_msg += f" Available: {available_options}. Attempting to continue."
                    logging.warning(log_msg)
            else:
                logging.info(f"Information (Process '{process_name_for_logging}'): Field 'severity' is empty. Skipping.")
        except (NoSuchElementException, ElementNotInteractableException, TimeoutException) as e:
            logging.error(f"Error (Process '{process_name_for_logging}'): Could not interact with 'severity' dropdown: {e}. Continuing to next field.", exc_info=True)
            submission_successful = False

        # 4. Responsible Person
        try:
            if responsible_person is not None:
                driver.find_element(By.ID, 'res').send_keys(str(responsible_person)) 
            else:
                logging.info(f"Information (Process '{process_name_for_logging}'): Field 'responsible_person' is empty. Skipping.")
        except (NoSuchElementException, ElementNotInteractableException, TimeoutException) as e:
            logging.error(f"Error (Process '{process_name_for_logging}'): Could not interact with 'responsible_person' field: {e}. Continuing to next field.", exc_info=True)
            submission_successful = False
            
        # 5. Commitment Date
        try:
            if commitment_date is not None and hasattr(commitment_date, 'strftime'):
                formatted_date = commitment_date.strftime('%d/%m/%Y')
                driver.find_element(By.ID, 'date').send_keys(formatted_date) 
            elif commitment_date is not None:
                logging.warning(f"Warning (Process '{process_name_for_logging}'): commitment_date ('{commitment_date}') is not a valid date object. Sending as text: '{str(commitment_date)}'.")
                driver.find_element(By.ID, 'date').send_keys(str(commitment_date))
            else:
                logging.info(f"Information (Process '{process_name_for_logging}'): Field 'commitment_date' is empty. Skipping.")
        except (NoSuchElementException, ElementNotInteractableException, TimeoutException) as e:
            logging.error(f"Error (Process '{process_name_for_logging}'): Could not interact with 'commitment_date' field: {e}. Continuing to next field.", exc_info=True)
            submission_successful = False

        # 6. Observation
        try:
            if observation is not None:
                driver.find_element(By.ID, 'obs').send_keys(str(observation)) 
            else:
                logging.info(f"Information (Process '{process_name_for_logging}'): Field 'observation' is empty. Skipping.")
        except (NoSuchElementException, ElementNotInteractableException, TimeoutException) as e:
            logging.error(f"Error (Process '{process_name_for_logging}'): Could not interact with 'observation' field: {e}. Continuing to next field.", exc_info=True)
            submission_successful = False

        # Click submit
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "submit"))).click() 
        except (TimeoutException, ElementNotInteractableException) as e:
            logging.error(f"Critical error (Process '{process_name_for_logging}'): Could not click 'submit' button: {e}. Cannot confirm submission.", exc_info=True)
            submission_successful = False 
            return 

        if not submission_successful:
            logging.warning(f"Warning (Process '{process_name_for_logging}'): Errors encountered while filling some fields. Submission may be incomplete or failed.")

        # Verify form submission
        try:
            alert_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "alert-success")))
            alert_text = alert_element.text.strip()

            if "Data sent" in alert_text:
                if "Queue ID " in alert_text:
                    try:
                        queue_id = alert_text.split("Queue ID ")[1].strip()
                        logging.info(f"Data for '{process_name_for_logging}' sent successfully. Queue ID: {queue_id}")
                    except IndexError:
                        logging.warning(f"Data for '{process_name_for_logging}' sent. 'Queue ID ' found, but ID could not be extracted. Full alert: '{alert_text}'")
                else: 
                    logging.warning(f"Data for '{process_name_for_logging}' sent, but Queue ID format not recognized. Full alert: '{alert_text}'")
            else: 
                logging.warning(f"Could not confirm submission for '{process_name_for_logging}' with 'Data sent'. Alert: {alert_text}")
        except TimeoutException:
            logging.error(f"Error (Process '{process_name_for_logging}'): Confirmation alert not found after submitting form.", exc_info=True)
        except Exception as e: 
            logging.error(f"Unexpected error during submission verification for '{process_name_for_logging}': {e}. Alert text (if available): '{alert_text if 'alert_text' in locals() else 'Not available'}'", exc_info=True)

    except Exception as e: 
        logging.error(f"General error during form submission for '{process_name_for_logging}': {e}", exc_info=True)

def send_status_email(process_name, status, observation, commitment_date, recipient_email):
    config = configparser.ConfigParser()
    try:
        if not config.read('config.ini'):
            logging.error("Error: 'config.ini' file not found. Please create one from 'config_example.ini' and fill in your credentials.")
            return 

        sender_email = config.get('SMTP', 'email')
        sender_password = config.get('SMTP', 'password')

        if not sender_email or not sender_password or sender_email == "tu_email@gmail.com" or sender_password == "tu_password_de_aplicacion":
            logging.error("Error: SMTP credentials not configured or are default values in 'config.ini'. Please update the file.") # 'tu_email@gmail.com' and 'tu_password_de_aplicacion' are from config_example.ini, should remain.
            return 

        email_subject = f'Process Status: {process_name}' # Translated
        date_string = commitment_date.strftime('%d/%m/%Y') if hasattr(commitment_date, 'strftime') else str(commitment_date)
        email_body = f"Process Status: {status}\nObservation: {observation}\nCommitment Date: {date_string}" # Translated
        
        email_message = MIMEMultipart()
        email_message['From'] = sender_email
        email_message['To'] = recipient_email
        email_message['Subject'] = email_subject
        email_message.attach(MIMEText(email_body, 'plain'))
        
        smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
        smtp_server.starttls()
        smtp_server.login(sender_email, sender_password)
        smtp_server.send_message(email_message)
        smtp_server.quit()
        
        logging.info(f"Email sent successfully to {recipient_email} for process '{process_name}'")
    except FileNotFoundError: 
        logging.error("Error: 'config.ini' file not found. Please create one from 'config_example.ini' and fill in your credentials.", exc_info=True)
    except (configparser.NoSectionError, configparser.NoOptionError) as e:
        logging.error(f"Configuration error in 'config.ini': {e}. Ensure [SMTP] section and 'email', 'password' options exist.", exc_info=True)
    except smtplib.SMTPAuthenticationError as e:
        logging.error(f"SMTP authentication error for process '{process_name}'. Verify credentials in 'config.ini'.", exc_info=True)
    except Exception as e:
        logging.error(f"Error sending email for process '{process_name}': {e}", exc_info=True)

if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(module)s - %(funcName)s - %(message)s',
        handlers=[logging.StreamHandler()]
    )
    process_excel_file()
    logging.info("Program has finished processing forms and sending emails.") # Translated
```
