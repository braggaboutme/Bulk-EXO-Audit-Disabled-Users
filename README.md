# Bulk-EXO-Audit-Disabled-Users
Gets and sets all users in the Azure tenant with Exchange mailboxes to have EXO auditing enabled and PowerShell remoting disabled.

Use at your own risk, this has been tested in environments with over 1 million users

In order to get this script to work, you'll need the correct permissions in your Azure tenant

# App Registration

1) Create an App Registration

2) Inside your app registration, click on "Certificates & secrets". Select "Upload Certificate" and upload your certificate.

**NOTE:** This certificate must be a non-CNG cert or else the script will fail.

If you are using self-signed certs, this script will help create a self signed cert to get you started.

_New-SelfSignedCertificate -FriendlyName "Exch Cert Auth" -Subject "Exch Cert Auth" -CertStoreLocation "cert:\CurrentUser\My" -KeySpec KeyExchange_

_$cert = Get-ChildItem -Path Cert:\CurrentUser\My\CERTTHUMBPRINTGOESHERE_

_Export-Certificate -Cert $cert -FilePath C:\Users\username\Desktop\exchauth.cer_

3) Next you'll need to expose the API's for this to work. Navigate to "API Permissions" and select "Add a permission". At the top, select "APIs my organization uses".
Type this number into the search "00000002-0000-0ff1-ce00-000000000000" to find the "Office 365 Exchange Online" APIs.
4) Select that API and select "Application permissions". Find "Exchange.ManageAsApp" from the list and select "Add permissions"
5) After the API is added, click the checkmark for "Grant admin consent" and select "yes"
6) Now that we have the correct API's, we need to assign Exchange permissions. Navigate to Azure Active Directory and go to "Roles and Administrators". Select "Exchange administrators" and click "Add assignment". Type the name of your App registration into the list and add that app to the Exchange Administrator Role.

**SECURITY WARNING:**
The reason this script needs the Exchange Administrator rather than Exchange recipient admin is because that AAD Role is required to set the Audit Enabled flag. If you don't want to set the Audit Enabled flag, then you can use other permissions like Exchange Recipient Admin.

