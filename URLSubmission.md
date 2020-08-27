### Running URL Submission API

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

## References
Public documentation the API: https://docs.microsoft.com/en-us/graph/api/resources/threatassessment-api-overview
