# AdminSubmissionsAPI

Admin Submission API allows submission of URLs, mail messages, file mail messages and files to Microsoft to re-scan and get newest verdict on submitted entity. Admin Submissions API is available both to Exchange Online Protection customers as well as to Office 365 ATP customers.
The repo provides a PowerShell script that allows to submit URLs and read re-scan result

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
10. Copy msal{AppID}://auth and paste the RedirectURI to script code as the value of $redirectURI variable. Click “Save” in the Azure AD app Authentication settings window.
11. On the App screen click “Overview” and copy “Application (client) ID” to the script code into the $clientID variable.
![AppID](/images/AppID.png)
12. Next, we need to assign user allowed to use this app. Assign user(s) to the app by following instruction from [this article](https://docs.microsoft.com/en-us/azure/active-directory/manage-apps/assign-user-or-group-access-portal#assign-users-or-groups-to-an-app-via-the-azure-portal) 
![Adding user](/images/AddUser.png)
13. Next in the Enterprise Application window, navigate to “Properties”. Select Yes next to “User assignment required” and click “Save”
![User assignment](/images/User_assignment_required.png)

## Script run parameters
* username (REQUIRED) – userPrincipalName of the user assigned to the app. 
* url – full URL (including protocol) to be submitted. This is parameter is used when single URL is to be submitted
* filepath - Path to the text file with URLs to be submitted. URLs should be in single column, without any header

## Examples
Bulk URL submission:
```.\AdminSubmissionAPI.ps1 -Username joe.doe@contoso.com -filepath URLs.txt```

Single URL submission:
```.\AdminSubmissionAPI.ps1 -Username joe.doe@contoso.com -url http://www.spamlink.contoso.com```

## License
We're completely open source and as matter of fact we also use some open source components in our report.
We use the following components in order to generate the report
•	Bootstrap, MIT License - https://getbootstrap.com/docs/4.0/about/license/
•	Fontawesome, CC BY 4.0 License - https://fontawesome.com/license/free


## References
Public documentation the API: https://docs.microsoft.com/en-us/graph/api/resources/threatassessment-api-overview
