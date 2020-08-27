## Using Email_AdminSubmissionAPI script
Email_AdminSubmissionAPI script searches for email to be submitted using Graph API. Then script either takes the found message and submits it or takes the attachment of the found message (the latter is useful when emails to be submitted are in the custom mailbox were they are reported by users using Cofense Outlook phishing button or Microsoft Report Message add-in). Because of that Azure AD app requires additional Microsoft Graph permissions:
* Read user mail (Mail.Read)
* Read user and shared mail (Mail.Read.Shared) (optional - needed if authenticated user should be able to submit emails not only for his/her mailbox but also from shared mailbox she/he has access to).
Script cannot run as a daemon. It requires authentication of the user performing the submission.

### Script search logic
The script can look for emails in single mailbox only. The mailbox to search is specified by *mailbox* parameter.
Script can look for email using Internet Message ID (if *InternetMessageID* parameter is used) or submit all emails received by the mailbox in x number of hours (specified by *agoHours* parameter). AgoHours and InternetMessageID attributes should be used in a mutually exclusive manner.

### Script parameters:
* username
* agoHours
* InternetMessageID
* mailbox
* attachment
* category

