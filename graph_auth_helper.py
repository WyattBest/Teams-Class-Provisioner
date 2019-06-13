import json
import msal

# This code based on Microsoft sample at https://github.com/AzureAD/microsoft-authentication-library-for-python/blob/dev/sample/confidential_client_secret_sample.py

# Load settings from disk
oauth_settings = json.load(open('settings.json'))['Microsoft']

# Create a preferably long-lived app instance which maintains a token cache.
app = msal.ConfidentialClientApplication(
    oauth_settings['application_id'], authority=oauth_settings['authority'],
    client_credential=oauth_settings['secret'],
)


def get_auth_header():
    """Returns token ready for use as 'Authorization' header. Checks memory cache before fetching a new token."""
    result = None

    # Check in-memory cache for existing token
    # Since we are looking for token for the current app, NOT for an end user,
    # we give account parameter as None.
    result = app.acquire_token_silent(oauth_settings['scope'], account=None)

    # If no token in cache, get a new one
    if not result:
        result = app.acquire_token_for_client(scopes=oauth_settings['scope'])

    if "access_token" in result:
        return result['token_type'] + ' ' + result['access_token']
    else:
        raise RuntimeError("Failed to get an authorization token.", result.get(
            "error"), result.get("error_description"), result.get("correlation_id"))
