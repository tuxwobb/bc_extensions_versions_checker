import datetime
from os import environ
import json

from dotenv import load_dotenv
from modules import BC_webservice, get_app_version, export_excel, export_csv, send_email

load_dotenv()

WS_URL = environ.get("WS_URL")
PROD_CLIENT_ID = environ.get("PROD_CLIENT_ID")
PROD_CLIENT_SECRET = environ.get("PROD_CLIENT_SECRET")
STAGE_CLIENT_ID = environ.get("STAGE_CLIENT_ID")
STAGE_CLIENT_SECRET = environ.get("STAGE_CLIENT_SECRET")
TENANT_ID = environ.get("TENANT_ID")
SCOPE = environ.get("SCOPE")
TOKEN_URL = environ.get("TOKEN_URL")
RECIPIENTS_FILENAME = environ.get("RECIPIENTS_FILENAME")
DATA_FILENAME = environ.get("DATA_FILENAME")

try:
    with open(DATA_FILENAME, "r", encoding="utf-8") as f:
        COUNTRIES = json.load(f)
except FileNotFoundError as e:
    print(f"File {DATA_FILENAME} not found.")
    exit(1)

result = []
for country in COUNTRIES:
    print(country["country"] + " AppSource")

    # Get version from AppSource
    for extension in country["extensions"]:
        if "appsource" in extension:
            version = get_app_version(extension["appsource"])
            print(extension["name"], version)
            result.append(
                {
                    "Country": country["country"],
                    "Environment": "AppSource",
                    "Extension": extension["name"],
                    "Version": version,
                }
            )
    print("")

    # Get version from Business Central
    for environment in country["environments"]:
        print(environment["url"])

        for extension in country["extensions"]:
            if environment["type"] == "prod":
                client_id = PROD_CLIENT_ID
                client_secret = PROD_CLIENT_SECRET
            else:
                client_id = STAGE_CLIENT_ID
                client_secret = STAGE_CLIENT_SECRET
            ws = BC_webservice(TENANT_ID, client_id, client_secret, SCOPE, TOKEN_URL)
            param = (
                "?$select=Name,Version,Is_Installed&$filter=Name eq '"
                + extension["name"]
                + "'"
            )
            response = ws.get_data(
                environment["url"], environment["company"], WS_URL, param
            )
            print(extension["name"])
            version = "N/A"
            for value in response["value"]:
                if value["Is_Installed"] == True:
                    version = value["Version"]
                    break
            print(version)
            result.append(
                {
                    "Country": country["country"],
                    "Environment": environment["type"],
                    "Extension": extension["name"],
                    "Version": version,
                }
            )
        print("")


# export results
date = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
filename = "extensions_versions_" + date + ".xlsx"
export_excel(result, filename)
# export_csv(result, "extensions_versions.csv")


# send email
try:
    with open(RECIPIENTS_FILENAME, "r", encoding="utf-8") as f:
        RECIPIENTS = json.load(f)
except FileNotFoundError as e:
    print(f"File {RECIPIENTS_FILENAME} not found.")
    exit(1)

SMTP_SERVER = environ.get("SMTP_SERVER")
SMTP_PORT = environ.get("SMTP_PORT")
EMAIL = environ.get("EMAIL")
PASSW = environ.get("PASSW")
SUBJECT = "BC Extension versions"
FILE = filename

for recipient in RECIPIENTS:
    CONTENT = f"Hi {recipient['name']},\n\nyou can find BC extensions instaled in stage and production environments + current version on AppSource in attachment.\n\nHave a nice day ahead!"
    send_email(
        smtp_server=SMTP_SERVER,
        smtp_port=SMTP_PORT,
        user=EMAIL,
        password=PASSW,
        email_from=EMAIL,
        email_to=recipient["email"],
        subject=SUBJECT,
        content=CONTENT,
        file=filename,
    )
