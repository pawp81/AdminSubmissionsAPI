# AdminSubmissionsAPI

Admin Submission API allows submission of URLs, mail messages, file mail messages and files to Microsoft to re-scan and get newest verdict on submitted entity. Admin Submissions API is available both to Exchange Online Protection customers as well as to Office 365 ATP customers.
The repo provides two PowerShell scripts:
1. for URLs submission
2. for email and emails from the attachment
Both scripts provide read of re-scan result capability.
The pre-requisites and preparation steps for URL and email submissions scripts related to the Azure AD app registration are the same and are described below,

## Pre-requisites
* Registered Azure AD app with Delegated permission: Read and write threat assessment requests (ThreatAssessment.ReadWrite.All). For creating new request, we need delegated permission to access users’ data as a signed-in user.
* Azure AD PowerShell: https://www.powershellgallery.com/packages/AzureAD/
* Azure AD user account. This user will be used to authenticate to Azure AD when running the script. The script uses Authorization Code flow OAUTH for authentication

## Deployment

### Azure AD app registration
1.  Navigate to the [Azure AD admin portal](https://aad.portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps)
2.  Click “New registration”
![App registration](/images/register.png)
3.  Enter name of your app for example "Threat Assessment". Leave “Accounts in this organizational directory only” option selected
4.  Select “public client/native” and click "Register"
5.  Click “API permissions” from left navigation menu.
6.  Click “Add a permission”. Click: "Microsoft Graph"
![API permissions](/images/APIpermissions.png)
7.  Click "Delegated permissions". Scroll down through the list of permission. Select "ThreatAssessment.ReadWrite.All". Click “Add permissions”.

![Permissions](/images/ThreatAssessment.ReadWrite.All.png)

8.  Refresh the list of permissions. Click “Grant admin consent for <your organization’s name>”. Click Yes.
![GrantConsent](/images/GrantConsent.png)
9.  Next click on “Authentication” from left navigation menu, click on “Switch to the old experience”
Select the checkbox next to the "msal{AppID}://auth (MSAL only)".
![Authentication](/images/authentication.png)
10. Copy msal{AppID}://auth value and paste it to script code as the value of $redirectURI variable. Click “Save” in the Azure AD app Authentication settings window.
11. On the App screen click “Overview” and copy “Application (client) ID” to the script code into the $clientID variable.
![AppID](/images/AppID.png)
12. Next, we need to assign user allowed to use this app. Assign user(s) to the app by following instruction from [this article](https://docs.microsoft.com/en-us/azure/active-directory/manage-apps/assign-user-or-group-access-portal#assign-users-or-groups-to-an-app-via-the-azure-portal) 
![Adding user](/images/AddUser.png)
13. Next in the Enterprise Application window, navigate to “Properties”. Select Yes next to “User assignment required” and click “Save”
![User assignment](/images/User_assignment_required.png)
14. In the script code update the path of the $authority url variable value with the default name of your tenant (for example: $authority="https://login.microsoftonline.com/contoso.onmicrosoft.com )

### Script operation instructions
After pre-requisites and deployment steps are fullfiled please read below manuals on how to execute the scripts:
* [URL submission](https://github.com/pawp81/AdminSubmissionsAPI/blob/master/URLSubmission.md)
* [Email submission](https://github.com/pawp81/AdminSubmissionsAPI/blob/master/EmailSubmission.md)
