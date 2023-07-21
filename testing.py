import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium.common.exceptions import TimeoutException
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import base64
from PIL import Image

def authenticate_gmail():
    try:
        with open('GmailCredentials.json') as file:
            gmail_credentials = json.load(file)
            gmail_email = gmail_credentials['email']
            gmail_password = gmail_credentials['otp']
            return gmail_email, gmail_password
    except FileNotFoundError:
        print("Gmail credentials JSON file not found.")
    except Exception as e:
        print(f"An error occurred during Gmail authentication: {str(e)}")
    return None, None

def authenticate_google_sheets(credentials_file):
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
        client = gspread.authorize(creds)
        print("Google Sheets authentication successful!")
        return client
    except FileNotFoundError:
        print("Google Sheets credentials JSON file not found.")
    except Exception as e:
        print(f"An error occurred during Google Sheets authentication: {str(e)}")
    return None

def open_google_sheet(client, spreadsheet_url):
    try:
        spreadsheet = client.open_by_url(spreadsheet_url)
        print("Google Sheet opened successfully!")
        return spreadsheet
    except gspread.exceptions.APIError as e:
        print(f"Error opening Google Sheet: {e.response}")
    except Exception as e:
        print(f"An error occurred while opening Google Sheet: {e}")
    return None

def get_current_datetime():
    current_date = datetime.datetime.now().date()
    current_time = datetime.datetime.now().time()
    return current_date, current_time

def process_email_data(row, current_date, current_time, email_content_tab, driver, gmail_email, gmail_password, html_bodies, recipients_list, subject_list, email_inputs_tab):
    email_id, company, vcp, meeting_name, subject, recipients, send_date_str, send_time_str, has_been_sent_today = row[:9]
    
    if not send_date_str:
        print(f"Send Date is empty for {email_id}. Skipping this row.")
        return

    send_date_str = send_date_str.strip() if send_date_str else "01/01/1970"
    send_time_str = send_time_str.strip() if send_time_str else "00:00"
    
    try:
        send_date = datetime.datetime.strptime(send_date_str, "%m/%d/%Y").date()
        send_time = datetime.datetime.strptime(send_time_str, "%H:%M").time()
    except Exception as e:
        print(f"Error processing date or time for {email_id}: {e}")
        print(f"Row data: {row}")
        return

    if send_date == current_date and send_time <= current_time and not has_been_sent_today:
        print(f"Processing email for {email_id}...")

        email_content_data = email_content_tab.get_all_values()
        content_header = email_content_data[0]
        content_data = email_content_data[1:]

        html_body = f"""
        <html>
            <body>
                <h2>Hi,</h2>
                <p>Please find the screenshot images from the following Tableau workbooks:</p>
        """

        tableau_links = {}
        for content_row in content_data:
            content_email_id, content_title, tableau_link, workbook_name, view_name = content_row[:5]

            if content_email_id == email_id:
                if content_email_id not in tableau_links:
                    tableau_links[content_email_id] = []
                tableau_links[content_email_id].append((tableau_link, workbook_name, view_name))

        captured_images = capture_tableau_screenshots(tableau_links, driver)

        for image in captured_images:
            workbook_name, view_name, image_path = image
            cropped_image_path = crop_image(image_path)
            image_data = open(cropped_image_path, 'rb').read()
            encoded_image = base64.b64encode(image_data).decode('utf-8')
            embedded_image = f'<img src="data:image/png;base64,{encoded_image}" alt="{workbook_name}_{view_name}">'
            html_body += f"<p>{view_name}</p>"
            html_body += embedded_image

        html_body += "</body></html>"

        html_bodies.append(html_body)
        recipients_list.append(recipients)
        subject_list.append(subject)

        print(f"Email processed for {email_id}")

        # Update 'Has been sent today?' to True for the processed row
        row_index = email_inputs_data.index(row) + 2  # Adjusted row index
        email_inputs_tab.update_cell(row_index, 9, 'True') # No need for calling 'update_has_been_sent_today()'

def login_gmail(driver):
    try:
        with open('GmailCredentials.json') as file:
            gmail_credentials = json.load(file)
            gmail_email = gmail_credentials['email']
            gmail_password = gmail_credentials['password']
            quoted_password = json.dumps(gmail_password)
            driver.get('https://accounts.google.com/')
            email_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'identifierId')))
            email_input.send_keys(gmail_email)
            email_input.send_keys(Keys.RETURN)
            time.sleep(5)
            password_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@type="password"]')))
            driver.execute_script(f"arguments[0].setAttribute('value', {quoted_password})", password_input)
            password_input.send_keys(Keys.RETURN)
            time.sleep(10)
    except FileNotFoundError:
        print("Gmail credentials JSON file not found.")
    except Exception as e:
        print(f"An error occurred during Gmail authentication: {str(e)}")
    return None, None

def check_login_tableau(driver):
    driver.get("https://10ay.online.tableau.com/#/site/turnriver/explore")
    try:
        # add wait condition here
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[text()='Explore']"))
        )
        print("Still logged")
    except TimeoutException:
        print("Relogging")
        login_tableau(driver)



def login_tableau(driver):
    jsonfile = 'TableauAuthP.json'
    with open(jsonfile, 'r') as file:
        credentials_tableau = json.load(file)
    options = webdriver.ChromeOptions()
    driver.get(credentials_tableau['server_url'])
    username_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input#email')))
    username_input.send_keys(credentials_tableau['username'])
    username_input.send_keys(Keys.RETURN)
    time.sleep(10)

def capture_tableau_screenshots(tableau_links, driver):
    images = []

    for email_id, tableau_link_list in tableau_links.items():
        for i, (tableau_link, workbook_name, view_name) in enumerate(tableau_link_list):
            if workbook_name and view_name:
                # Navigate to the tableau link
                driver.get(tableau_link)
                time.sleep(10)

                # Capture screenshot as image file
                screenshot_file = f"{workbook_name}_{view_name}.png"
                driver.save_screenshot(screenshot_file)

                images.append((workbook_name, view_name, screenshot_file))

    return images

def crop_image(image_path):
    image = Image.open(image_path)
    image_data = image.load()

    # Remove the background color
    width, height = image.size
    for x in range(width):
        for y in range(height):
            r, g, b, a = image_data[x, y]
            if r == 14 and g == 14 and b == 14:
                image_data[x, y] = (0, 0, 0, 0)

    # Crop the image to remove any transparent pixels
    bbox = image.getbbox()
    cropped_image = image.crop(bbox)

    # Save the cropped image to a file
    cropped_image_path = "cropped_image.png"
    cropped_image.save(cropped_image_path)

    return cropped_image_path

def send_email(sender_email, sender_password, recipients, subject, html_bodies):
    for i in range(len(html_bodies)):
        html_body = html_bodies[i]
        recipient = recipients[i]
        email_subject = subject[i]

        # Create the root message container
        message_root = MIMEMultipart("related")
        message_root["From"] = sender_email
        message_root["To"] = recipient
        message_root["Subject"] = email_subject

        # Create the HTML message
        message_html = MIMEText(html_body, "html")

        # Attach the HTML message to the root message container
        message_root.attach(message_html)

        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.ehlo()
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient.split(","), message_root.as_string())

        print(f"Email sent to {recipient} with subject '{email_subject}'")

def main():
    global email_inputs_data
    global spreadsheet
    gmail_email, gmail_password = authenticate_gmail()
    if not gmail_email or not gmail_password:
        return

    credentials_file = 'TR.json'
    client = authenticate_google_sheets(credentials_file)
    if not client:
        return

    spreadsheet_url = 'https://docs.google.com/spreadsheets/d/example'
    spreadsheet = open_google_sheet(client, spreadsheet_url)
    if not spreadsheet:
        return

    email_inputs_tab = spreadsheet.worksheet("Email Inputs")
    email_content_tab = spreadsheet.worksheet("Email Content")

    current_date, current_time = get_current_datetime()

    email_inputs_data = email_inputs_tab.get_all_values()
    header = email_inputs_data[0]
    data = email_inputs_data[1:]

    driver_path = 'pathto\chromedriver.exe'
    service = Service(driver_path)
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=service, options=options)
    login_gmail(driver)
    login_tableau(driver)

    while True:
        html_bodies = []
        recipients_list = []
        subject_list = []

        for row in data:
            process_email_data(row, current_date, current_time, email_content_tab, driver, gmail_email, gmail_password, html_bodies, recipients_list, subject_list, email_inputs_tab, email_inputs_data)

        if html_bodies:
            send_email(gmail_email, gmail_password, recipients_list, subject_list, html_bodies)

        print("Task completed successfully!")
        time.sleep(60)
        check_login_tableau(driver)

if __name__ == "__main__":
    main()

