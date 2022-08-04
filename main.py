# Office 365 Automatic Email Processor
# Written by Adrian Arnett 8/3/22

import requests
import configparser
import msal


class AuthenticationError(Exception):
    """
    Generic authentication exception class
    """
    pass


# read oauth2 token from config file
config = configparser.ConfigParser()
config.read("./settings.cfg")
AZURE_DOMAIN = config["azure"]["domain"]
AZURE_TENANT_ID = config["azure"]["tenant_id"]
MS_IDENTITY_APP_ID = config["microsoft_identity"]["app_id"]
MS_IDENTITY_APP_SECRET_ID = config["microsoft_identity"]["secret_id"]
MS_IDENTITY_APP_SECRET_VALUE = config["microsoft_identity"]["secret_value"]
MS_IDENTITY_SCOPE = config["microsoft_identity"]["scope"]


def authenticate():
    app = msal.ConfidentialClientApplication(
        MS_IDENTITY_APP_ID, client_credential=MS_IDENTITY_APP_SECRET_VALUE,
        authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}")
    response = app.acquire_token_for_client([MS_IDENTITY_SCOPE])

    if "error" in response:
        raise AuthenticationError(f"Invalid authentication response:\n{response['error_description']}")

    print("successfully authenticated with MSAL")
    return response


key = authenticate()["access_token"]
