# Office 365 Automatic Email Processor
# Written by Adrian Arnett 8/3/22

import requests
import sched
import time
import configparser

# libs
import msal


class AuthenticationError(Exception):
    """
    Generic authentication exception class
    """
    pass


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


def obtain_all_emails():  # TODO: Fix this to use a mail enabled security group
    response = requests.get("https://graph.microsoft.com/v1.0/users/83eed5d0-22aa-4ee3-8d9a-0f5408f9d365/messages",
                            headers={"Authorization": f"Bearer {AUTH_KEY}"})
    return response.json()


def obtain_email_delta():
    response = requests.get(
        "https://graph.microsoft.com/v1.0/users/83eed5d0-22aa-4ee3-8d9a-0f5408f9d365/mailFolders/Inbox/messages/delta",
        headers={"Authorization": f"Bearer {AUTH_KEY}"})
    return response.json()


AUTH_KEY = authenticate()["access_token"]

while True:
    print("checking authentication token...")
    scheduler.run(False)  # check for scheduled events - non-blocking
    print("checking for new emails...")
    print(obtain_email_delta())
    print("end check")
    time.sleep(APPLICATION_SLEEP_TIME)

