# Macro for Microsoft Graph OAuth

## Background
These macros are for a Cisco video device that provides a means of utilizing OAuth tokens for authentication to Microsoft Graph API's.

> NOTE: Because the clientID and clientSecret are stored in a macro on the device, precautions need to be made.
> Any person who has read-only access to the video device or Control Hub, will be able to view the clientId and clientSecret.
> As such, you can look to leverage a centralised service where these values are protected. Alternatively ensure the OAuth scopes are limited as not to cause any problems if compromised.

The macro's are broken into 3 seperate files:
1. MsftListUsers.js - this file contains the standard API call Microsoft Graph API service
2. MsftManageTokens.js - this file manages the existing token and if needed will refresh the token
3. MsftSavedTokens.js - this file stores the current access_token and expiry value

## Setup
In order to obtain a clientId and clientSecret, login to [Microsoft Entra](https://entra.microsoft.com/). Select "App registrations" and then "New registration". Provide a name for the registration and then select the "Register" button.

Copy the "Application (client) ID" and "Directory (tenant) ID" values to notepad. Then select "Certificates & secrets" from the left menu. On the "Clients secret" tab, select the "New client secret" button. Provide a name for the secret and select the duration for which the secret will expire. Once you select the "Add" button, a client secret is generated. Copy the "Value" field to notepad.

Upload the three macro files (.js) from this repo to your video device. Ensure they are saved locally but at this stage do not enable them.

In the MsftManageTokens macro, update the clientId, clientSecret and tenantId to the values you obtained earlier.

For the MsftListUsers macro,  enable _only_ this macro (the other two do not need to be enabled). You should see a list of all users in the Org in the console of the Macro editor.

If you exit the Macro Editor page (top left) and then return - opening the MsftSavedTokens macro should now have a new access_token and expires_at value.
