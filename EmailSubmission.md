## General information about Email_AdminSubmissionAPI script
[Email_AdminSubmissionAPI](https://github.com/pawp81/AdminSubmissionsAPI/edit/master/Email_AdminSubmissionAPI.ps1) script searches for email to be submitted using Graph API. After finding it either takes the message and submits it or takes the attachment of the found message and submits it (the latter is useful when emails to be submitted are in the custom mailbox were they are reported by users using Cofense Outlook phishing button or Microsoft Report Message add-in). In both cases it uses Admin Submission API for submission. Because of search capability Azure AD app requires Microsoft Graph permissions:
* Read user mail (Mail.Read)
* Read user and shared mail (Mail.Read.Shared) (optional - needed if authenticated user should be able to submit emails not only for his/her mailbox but also from shared mailbox she/he has access to).
Script can run as a daemon. However, it requires authentication of the user performing the submission. With usage of MSAL PS PowerShell module, refresh token can be saved and reused to obtain access token. Therefore no logon prompts are expected after initial run of the script.

![User_Consent](/images/User_consent.png)
If organization doesn't allow users consent (as shown on above screenshot), admin will need to consent Mail.Read and Mail.Read.Shared permissions.

### Script search logic
The script can look for emails in the single mailbox only. The mailbox to search is specified by *mailbox* parameter.
Script can look for email using Internet Message ID (if *InternetMessageID* parameter is used) or submit all emails received by the mailbox in x number of hours since now (specified by *agoHours* parameter). AgoHours and InternetMessageID attributes should be used in a mutually exclusive manner.

### Script parameters:
* *username* - (mandatory)userPrincipalName of the user assigned to the app
* *agoHours* - Specifies how many hours from now, should the script look for emails to submit. For example setting -agoHours 24 will look for all emails received by the mailbox in last 24h and submit them (or their attachments)
* *agoMinutes* - Specifies how many minutes from now, should the script look for emails to submit. For example setting -agoMinutes 15 will look for all emails received by the mailbox in last 15 minutes and submit them (or their attachments)
* *InternetMessageID* - InternetMessageID of the message to be submitted or its attachment. Should be provided without <> brackets
* *mailbox* - (mandatory) Mailbox from which email is to be submitted
* *attachment* - If set to true specifies that attachment of the found message should be submitted.  If not set message itself specified by InternetMessageID or by agoHours will be submitted.
* *category* - (mandatory) Specifies what category of the email submission is used. Allowed values: "phishing", "spam".
* *confirm* - should be set to $true if you want to manually confirm every submission

### References
* mailAssessmentRequest resource type https://docs.microsoft.com/en-us/graph/api/resources/mailassessmentrequest?view=graph-rest-1.0
* emailFileAssessmentRequest resource type https://docs.microsoft.com/en-us/graph/api/resources/emailfileassessmentrequest?view=graph-rest-1.0
* Create a mail assessment request https://docs.microsoft.com/en-us/graph/api/informationprotection-post-threatassessmentrequests?view=graph-rest-1.0&tabs=http#examples


