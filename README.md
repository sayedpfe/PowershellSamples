# Introduction
This repository hosts samples for use with the Microsoft Graph PowerShell SDK.

---------------------------------------

## Setup

To leverage the Microsoft Graph API, we will need to setup and configure some prerequisites.
* Self signed certificate for secure authentication.
* Azure Active Directory application creation and configuration.

---------------------------------------

### Create a Client Certificate

In order to authenticate without a user present, we will use a self signed client certificate.
Before running the following cmdlet, be sure to change "**`<Subject>`**" to something that uniquely identifies its purpose. i.e., **GraphPOC.PS-msgraph**.

1. Using PowerShell on your developer computer, run the the following to create your client certificate. It will be stored in your users' certificate store.

    ```
    $cert = New-SelfSignedCertificate `
            -CertStoreLocation cert:\currentuser\my `
            -Subject <Subject> `
            -KeyDescription "Used to access Microsoft Graph API" `
            -NotAfter (Get-Date).AddYears(2)
    ```

2. Copy the thumbprint and edit the **.\config\clientconfiguration.json** file by replacing the appropriate property value (Thumbprint).
    ```
    $cert.Thumbprint | clip
    ```
3. Export the certificate to a secure location for later use.
    ```
    Export-Certificate -Type CERT -Cert $cert -FilePath c:\temp\graphapi.cer;
    ```
More details on using certificates can be found here: https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-certificate-credentials.

---------------------------------------

### Create an Azure AD Application

1. Sign in to your Azure Account through the Azure portal (https://portal.azure.com/).
2. Select **Azure Active Directory**.
3. Select **App registrations**.
4. Select **New registration**.
5. Name the application. I use the subject from the client certificate for easy identification. i.e., **GraphPOC.PS-msgraph**
6. Under "**Supported account types**", ensure the default, "**Accounts in this organizational directory only (contoso only - Single tenant)**", is selected and click "**Register**". You will be redirected to the app configuration page.
7. In the **Overview**, copy the "**Application (client) ID**" and "**Directory (tenant) ID**" and edit the **.\config\clientconfiguration.json** file by replacing the appropriate property values (TenantId and ClientId).
8. Cick on "**Authentication**".
9. Under **"Platform Configurations"** click on **"Add a Platform"**. Click on the tile that reads "**Mobile and desktop applications**". Then, check the first box in front of https://login.microsoftonline.com/common/oauth2/nativeclient and click "**Configure**".
10. Move to "**API permissions**" and select "**Add a permissions**" and click on "**Microsoft Graph**"
11. Next, choose "**Application permissions**". Select the following and click "**Add permissions**".

    ```
    User.ReadWrite.All
    Directory.Read.All
    Calendars.Read
    Mail.Read
    Mail.Send
    AuditLog.Read.All
    Reports.Read.All
    SecurityEvents.Read.All
    Group.ReadWrite.All
    AuditLog.Read.All
    Sites.Read.All
    ```

12. Click "**Grant admin consent for contoso**" and then click "**Yes**" to the prompt.
13. Finally, navigate to "**Certificates & secrets**" and click the "**Upload certificate**" button, browse to where you exported your certificate and select it. For security reasons, you should delete the certificate file you exported after it's been uploaded. 
    
More details on the App Registration process can be found here: https://docs.microsoft.com/en-us/graph/auth-register-app-v2

---------------------------------------

## Usage

Open PowerShell and install the Microsoft.Graph.Authentication module.

```powershell
Install-Module -Name Microsoft.Graph.Authentication -CurrentUser
```

Navigate to the projects root folder, then import your configuration and connect to Graph.

```powershell
$config = Get-Content .\config\clientconfiguration.json -Raw | ConvertFrom-Json

Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint
```

Using the samples simply require that you've performed the setup steps and are connected to Graph.

```powershell
$user = 'user@contoso.com'
Invoke-MgGraphRequest -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$user"

Name              Value
----              -----
@odata.context    https://graph.microsoft.com/v1.0/$metadata#users/$entity
businessPhones    {}
displayName       brandon
givenName         brandon
jobTitle
mail              user@contoso.com
mobilePhone
officeLocation
preferredLanguage
surname
userPrincipalName user@contoso.com
id                48d9c121-bb2c-402b-bedf-612296500d2e
```