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
The creation of the clientId and clientSecret *** Add Graph specific material ***

Upload the three macro files (.js) from this repo to your video device. Ensure they are saved locally but at this stage do not enable them.

In the MsftManageTokens macro, update the clientId, clientSecret and tenantId to the values you obtained earlier.

For the MsftListUsers macro,  enable _only_ this macro (the other two do not need to be enabled). You should see a list of all users in the Org in the console of the Macro editor.

If you exit the Macro Editor page (top left) and then return - opening the MsftSavedTokens macro should now have a new access_token and expires_at value.
