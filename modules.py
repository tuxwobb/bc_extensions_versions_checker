import csv
import smtplib
from email.message import EmailMessage

import requests
import pandas

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# class for calling Business Central web services
class BC_webservice:
    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        scope: str,
        token_url: str,
    ):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.scope = scope
        self.token_url = token_url

    def __get_token(self):
        data = {
            "grant_type": "client_credentials",
            "CLIENT_ID": self.client_id,
            "CLIENT_SECRET": self.client_secret,
            "SCOPE": self.scope,
        }
        try:
            response = requests.post(self.token_url, data=data, timeout=30)
        except Exception as e:
            print(f"Error occurred while fetching token.")
            exit(1)

        if response.status_code != 200:
            print("Failed to obtain token.")
            exit(1)

        return response.json().get("access_token")

    def get_data(
        self,
        environment_name: str,
        company_name: str,
        webservice_name: str,
        parameters: str,
    ):
        token = self.__get_token()
        service_url = f"https://api.businesscentral.dynamics.com/v2.0/{
            environment_name}/ODataV4/Company('{company_name}')/{webservice_name}{parameters}"
        headers = {
            "Authorization": f"Bearer {
            token}",
            "Accept": "application/json",
        }
        try:
            response = requests.get(service_url, headers=headers, timeout=30)
        except Exception as e:
            print(f"Error occurred while fetching data: {e}")
            exit(1)

        if response.status_code == 500:
            print(
                f"Server error when accessing {service_url} - check the service and parameters: {response.text}"
            )

        if response.status_code == 401:
            print(
                f"Unauthorized access {service_url} - check your credentials and permissions: {response.text}"
            )

        if response.status_code == 404:
            print(
                f"Resource not found: {service_url} - check the service and parameters: {response.text}"
            )

        if response.status_code == 200:
            return response.json()

        return {"value": [{"Name": "N/A", "Version": "N/A", "Is_Installed": False}]}


# Get version from AppSource
def get_app_version(url: str):
    service = Service(ChromeDriverManager().install())

    options = webdriver.ChromeOptions()
    options.add_argument("--headless")

    driver = webdriver.Chrome(service=service, options=options)
    try:
        driver.get(url)
        wait = WebDriverWait(driver, 10)
        version_label = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    '//*[@id="pdpTabs"]/div[2]/div/div[1]/div/div[1]/div/div[3]/span',
                )
            )
        )
        version_value = version_label.find_element(
            By.XPATH,
            '//*[@id="pdpTabs"]/div[2]/div/div[1]/div/div[1]/div/div[3]/span',
        )
        return version_value.text
    except Exception as e:
        print(f"Error occurred while fetching version: {e}")
        return "Error"
    finally:
        driver.quit()


# Export results to excel
def export_excel(result: dict, excel_file: str):

    df = pandas.DataFrame(result)
    pivot_df = df.pivot_table(
        index=["Country", "Extension"],  # lines
        columns="Environment",  # columns
        values="Version",  # values
        aggfunc="first",  # if duplicity, take the first one
    ).reset_index()
    pivot_df.to_excel(excel_file, index=False)
    return pivot_df


# Export results to csv
def export_csv(result: dict, csv_file: str):
    with open(csv_file, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=["country", "environment", "extension", "version"]
        )

        writer.writeheader()
        writer.writerows(result)


# Send email with attachment
def send_email(
    smtp_server: str,
    smtp_port: int,
    user: str,
    password: str,
    email_from: str,
    email_to: str,
    subject: str,
    content: str,
    file: str,
):

    msg = EmailMessage()
    msg["From"] = email_from
    msg["To"] = email_to
    msg["Subject"] = subject
    msg.set_content(content)

    try:
        with open(file, "rb") as f:
            msg.add_attachment(
                f.read(), maintype="application", subtype="xlsx", filename=file
            )
    except FileNotFoundError as e:
        print(f"File {file} not found.")
        exit(1)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(user, password)
            server.send_message(msg)

        print("Email sent to " + email_to + " (" + file + ")")
    except Exception as e:
        print(f"Error occurred while sending email: {e}")
        exit(1)
