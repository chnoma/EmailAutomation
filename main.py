# Office 365 Automatic Email Processor
# Written by Adrian Arnett 8/3/22

import requests
import sched
import datetime
import time
import configparser
import urllib.parse

# libs
import msal


class AuthenticationError(Exception):
    """
    Generic authentication exception class
    """
    pass


APPLICATION_SEPARATOR = "----------------------------------------------------------------"
APPLICATION_REAUTH_TIME = 0.8  # the factor to multiply expires_in to renew the authentication token - provides margin for error
APPLICATION_SLEEP_TIME = 15  # check for new emails every 20 seconds
# read oauth2 token from config file
config = configparser.ConfigParser()
config.read("./settings.cfg")
AZURE_DOMAIN = config["azure"]["domain"]
AZURE_TENANT_ID = config["azure"]["tenant_id"]
MS_IDENTITY_APP_ID = config["microsoft_identity"]["app_id"]
MS_IDENTITY_APP_SECRET_ID = config["microsoft_identity"]["secret_id"]
MS_IDENTITY_APP_SECRET_VALUE = config["microsoft_identity"]["secret_value"]
MS_IDENTITY_SCOPE = config["microsoft_identity"]["scope"]
APPLICATION_MAILBOX = f'{config["application"]["mailbox"]}@{AZURE_DOMAIN}'

scheduler = sched.scheduler()  # init scheduler
app_delta_link = "https://graph.microsoft.com/v1.0/users/83eed5d0-22aa-4ee3-8d9a-0f5408f9d365/mailFolders/Inbox/messages/delta"


def authenticate():
    print("authentication required")
    app = msal.ConfidentialClientApplication(
        MS_IDENTITY_APP_ID, client_credential=MS_IDENTITY_APP_SECRET_VALUE,
        authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}")
    response = app.acquire_token_for_client([MS_IDENTITY_SCOPE])
    expiration_time = response["expires_in"]
    renewal_time = expiration_time * APPLICATION_REAUTH_TIME
    if "error" in response:
        raise AuthenticationError(f"Failed to authenticate:\n{response['error_description']}")
    scheduler.enter(renewal_time, 1, authenticate)
    print("successfully authenticated with MSAL")
    print(f"re-authentication in {renewal_time} seconds")

    return response


def request_get_authenticated(url):
    response = requests.get(url, headers={"Authorization": f"Bearer {AUTH_KEY}"})
    response.raise_for_status()
    return response.json()


def obtain_all_emails():  # TODO: Fix this to use a mail enabled security group
    return request_get_authenticated(
        "https://graph.microsoft.com/v1.0/users/83eed5d0-22aa-4ee3-8d9a-0f5408f9d365/messages")


def check_new_emails():
    global app_delta_link
    response = request_get_authenticated(app_delta_link)  # TODO: handle @odata.nextLink (exists if pages exceed limit)
    email_count = len(response['value'])
    app_delta_link = response['@odata.deltaLink']
    return email_count


print(APPLICATION_SEPARATOR)
print(str(datetime.datetime.now()))
print("starting application...")
AUTH_KEY = authenticate()["access_token"]
check_new_emails()  # ignore pre-existing emails

while True:
    print(APPLICATION_SEPARATOR)
    print(str(datetime.datetime.now()))
    print("checking authentication token...")
    scheduler.run(False)  # check for scheduled events - non-blocking
    print("checking for new emails...")
    print(f"{check_new_emails()} new emails")  # TODO: Seems like an email can persist in some conditions?
    print("end check")
    time.sleep(APPLICATION_SLEEP_TIME)
