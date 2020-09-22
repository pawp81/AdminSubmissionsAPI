  
<#
	.SYNOPSIS
		The Admin Submission API
	.DESCRIPTION
       
	.NOTES
		Pawel Partyka
		Senior Program Manager - Microsoft
        ppartyka@microsoft.com
        
        ############################################################################
        This sample script is not supported under any Microsoft standard support program or service. 
        This sample script is provided AS IS without warranty of any kind. 
        Microsoft further disclaims all implied warranties including, without limitation, any implied 
        warranties of merchantability or of fitness for a particular purpose. The entire risk arising 
        out of the use or performance of the sample script and documentation remains with you. In no
        event shall Microsoft, its authors, or anyone else involved in the creation, production, or 
        delivery of the scripts be liable for any damages whatsoever (including, without limitation, 
        damages for loss of business profits, business interruption, loss of business information, 
        or other pecuniary loss) arising out of the use of or inability to use the sample script or
        documentation, even if Microsoft has been advised of the possibility of such damages.
        ############################################################################    
	.LINK
        about_functions_advanced


Parameters
Find an email by InternetMessageID in the mailbox and submit it.
Don't use brackets when providing InternetMessageID
Set attachment: $true is email to be submitted is in the attachment.

Examples:
 Submit single email received by John Doe. Refer the message by its InternetmessageID 
.\Email_AdminSubmissionAPI.ps1 -Username john.doe@contoso.com  -InternetMessageID MWHPR01MB2574F219BE7EDC5E41153730CA290@MWHPR01MB2574.prod.exchangelabs.com -mailbox john.doe@contoso.com -category phishing

 Submit single email received by SOC mailbox. Refer the message by its InternetmessageID. Instead of submitting actual message script submits its attachment
.\Email_AdminSubmissionAPI.ps1 -Username admin@contoso.com  -InternetMessageID MWHPR01MB2574F219BE7EDC5E41153730CA290@MWHPR01MB2574.prod.exchangelabs.com -mailbox soc@contoso.com -category phishing -attachment $true

 Submit attachments of the emails received by soc@contoso.com mailbox in last 24h. Submit emails as phishing. Use admin@contoso.com credentials for authentication
.\Email_AdminSubmissionAPI.ps1 -Username admin@contoso.com -agoHours 24 -mailbox soc@contoso.com -category phishing -attachment $true

 Submit emails received by John Doe mailbox in last 24h as phishing.
.\Email_AdminSubmissionAPI.ps1 -Username admin@contoso.com  -agoHours 24 -mailbox john.doe@contoso.com -category phishing

#>

param(
        [Parameter(Mandatory = $true)]
        [string]$Username,
		[int]$agoHours,
		[string]$InternetMessageID,
		[Parameter(Mandatory = $true)]
		[string]$mailbox,
		[boolean]$attachment,
		[Parameter(Mandatory = $true)]
		[string]$category, #allowed values: "phishing", "spam"
		[boolean]$confirm
    )

# configure below variables according to: https://github.com/pawp81/AdminSubmissionsAPI
$clientId =  ""
$redirectUri = "msal://auth" 
$authority = "https://login.microsoftonline.com/<tenantname>"

# List of static variables.
$resourceURI = "https://graph.microsoft.com"
$GraphUrl="https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests"
$day=(get-date).day
$month=(get-date).month
$year=(get-date).year
$random=get-random -Maximum 10000
$fileSubmissionIDs="SubmissionIDs-list-$day-$month-$year_$random.txt"

function Get-AccessToken{
	
	try
	{
		if(get-childItem accessToken.txt)
		{
			$tryAccessToken=get-childItem accessToken.txt
			[datetime]$LastWriteTime=$tryAccessToken.LastWriteTime
			$TokenExpiration=$LastWriteTime.addhours(1)
			$currentdate=get-date
		}
	}
	catch
	{
	}
	if ($TokenExpiration -le $currentdate)
	{
		write-host "New Access Token needed"
		try {
		$AadModule = Import-Module -Name AzureAD -ErrorAction Stop -PassThru
		}
		catch {
		try
			{
			$AadModule = Import-Module -Name AzureADPreview -ErrorAction Stop -PassThru
			}
			catch{
			throw 'Prerequisites not installed (AzureAD PowerShell module not installed)'
			}
		}
		$adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
		$adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
		[System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
		[System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
		$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
	
		# Get token by prompting login window.
		$platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Always"
		$userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($username, "OptionalDisplayableId")
		$authResult = $authContext.AcquireTokenAsync($resourceURI, $ClientID, $RedirectUri, $platformParameters)
		$accessToken = $authResult.result.AccessToken
		$accessToken > accessToken.txt
	
		$tokenPayload=$accessToken.Split(".")[1]
		while ($tokenPayload.Length % 4) { Write-Verbose "Invalid length for a Base-64 char array or string, adding ="; $tokenPayload += "=" }
		Write-Verbose "Base64 encoded (padded) payoad:"
		Write-Verbose $tokenPayload
		#Convert to Byte array
		$tokenByteArray = [System.Convert]::FromBase64String($tokenPayload)
		#Convert to string array
		$tokenArray = [System.Text.Encoding]::ASCII.GetString($tokenByteArray)
		Write-Verbose "Decoded array in JSON format:"
		Write-Verbose $tokenArray
		#Convert from JSON to PSObject
		$tokobj = $tokenArray | ConvertFrom-Json
		$tokobj
		$DecodedText = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($accessTokendecoded))	
	}
	else
		{
			$accessToken=get-content accessToken.txt
		}
		
	return $accessToken.ToString()
}

function Find-Email{
Param
	(	
		$InternetMessageID,
		$mailbox,
		$attachment
	)
	# This function searches for an email and returns Graph-based unique identifier of the message.
	[hashtable]$return = @{}
	
	$accessToken=Get-AccessToken
	$Headers= @{"Content-Type" = "application/json" ; "Authorization" = "Bearer " + $accessToken}
	if ($InternetMessageID)
	{
		$URI = "https://graph.microsoft.com/v1.0/users/$mailbox/mailFolders/inbox/messages?`$select=id,toRecipients&`$filter=InternetMessageID eq '<$InternetMessageID>'"
	}
	if ($agoHours)
	{
		$startdate=(get-date).AddHours(-$agohours).tostring("yyyy-MM-dd")
		$URI = "https://graph.microsoft.com/v1.0/users/$mailbox/mailFolders/inbox/messages?`$top=1000&`$select=id,toRecipients&`$filter=ReceivedDateTime ge $startdate"
	}
	
	write-host "Searching for:" $URI
	$MessageJSON=Invoke-WebRequest -Uri $URI -Headers $headers
	$Messages=$MessageJSON.content | ConvertFrom-JSON
	if ($Messages.value.length -eq 0)
	{
		write-host "Message not found. Existing" -foregroundcolor Red
		exit
	}
	$MessageIDs = @()
	foreach ($Message in $Messages)
	{
		#Saving all messageIDs matching search criteria to the array
		
		$MessageIDs+=$Message.value.id
	}

	
	if ($attachment -eq $true)
	{
		#Looking for attachment ID. I need to download it to the disk.
		$path = @()
		foreach ($MessageID in $MessageIDs)
		{
			$MessageAttachmentURI="https://graph.microsoft.com/v1.0/users/$mailbox/messages/$MessageID/attachments/"
			try{
				$MessageAttachmentJSON=Invoke-WebRequest -Uri $MessageAttachmentURI -Headers $headers
				$MessageAttachment= $MessageAttachmentJSON | ConvertFrom-JSON
			}
			catch
			{
				write-host "`n"
				write-host "Attachment for Message ID $MessageID not found" -foregroundcolor red
			}
			
			if ($MessageAttachment.value.id.length -gt 0)
			{
				write-host "Message attachment ID found: " $MessageAttachment.value.id -foregroundcolor green
				$AttachmentID=$MessageAttachment.value.id
				$AttachmentName=$MessageAttachment.value.name
							
				#downloading the attachment to current folder
				$temppath= $attachmentID+".eml"
				$AttachmentFetchURL="https://graph.microsoft.com/v1.0/users/$mailbox/messages/$MessageID/attachments/$AttachmentID/`$value"
				write-host "Sending request to fetch attachment: " $AttachmentFetchURL
				Invoke-WebRequest -Uri $AttachmentFetchURL -Headers $headers -outfile $temppath
			
				$path += $attachmentID+".eml"
			}
		}
		$return.path=$path
	}
	else
	{
	#Preparing MessageURI to be used in the submission request body
		if ($mailbox)
		{
			$MessageURI=@()
			foreach ($MessageID in $MessageIDs)
			{
				$MessageURI+="https://graph.microsoft.com/v1.0/users/$mailbox/messages/$MessageID"
			}
		}
		else
		{
			$UserGUID=$Message."@odata.context" |select-string -Pattern '[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}' | foreach {$_.matches.value}
			$MessageURI="https://graph.microsoft.com/v1.0/users/$UserGUID/messages/$MessageID"
		}	
	$return.MessageURI=$MessageURI
	}
	
	$return.recipient=$mailbox
	return $return
}

function Submit-Email {
# Function to submit Email from user mailbox
	Param
	(	
		$MessageURI,
		$recipient,
		$category,
		$attachmentpath
	)	
	$ThreatRequestIDs=@()
	
	foreach ($MessageURI_ID in $MessageURI)
	{
		write-host "`n"
		write-host "Evaluating " $MessageURI_ID -foregroundcolor green
		if ($confirm -eq $true)
		{
			$submit=read-host "Submit this Email? [YES]"
		}
		else
		{
			$submit="yes"	
		}
		if ($submit -eq "yes" -or $submit -eq "y")
		{
		$body = @"
		{
	"@odata.type": "#microsoft.graph.mailAssessmentRequest",
	"recipientEmail": "$recipient",
	"expectedAssessment": "block",
	"category": "$category",
	"messageUri": "$MessageURI_ID"
}
"@
			$accessToken=Get-AccessToken
			$Headers= @{"Content-Type" = "application/json" ; "Authorization" = "Bearer " + $accessToken}
			
			$error.clear()
			$ErrorOccured = $false
			try
			{
				$ThreatRequestJSON=Invoke-WebRequest -Uri $GraphURL -Headers $headers -Body $body -Method POST -ContentType 'application/json; charset=utf-8'
			}
			catch
			{
				write-host "Error submitting Email: " -foregroundcolor red
				write-host $_ -foregroundcolor red
				$ErrorOccured=$true
			}
			if (!$ErrorOccured)
			{
				$ThreatRequest=$ThreatRequestJSON.content |convertfrom-JSON
				write-host "Email submitted. Submission ID: " $ThreatRequest.id -foregroundcolor green
				write-host "`n"
				$ThreatRequestIDs+=$ThreatRequest.id
				$ThreatRequest.id | out-file $fileSubmissionIDs -append
				$anythingsubmitted=$true
			}
		$submit="no"
		}
	}
	if ($anythingsubmitted)
	{
		write-host "All request IDs:"
		$ThreatRequestIDs
		$checkSubmission=Read-Host "Do you want to check Submission status? [Yes\No]"
		if ($checkSubmission -eq "Yes")
		{
			Check-Submission -ThreatRequestIDs $ThreatRequestIDs
		}
	}
}

function Submit-Attachment {
# Function to submit email attachment
Param
(	
	$recipient,
	$category,
	$attachmentnames
)	
	$ThreatRequestIDs=@()
	foreach ($attachmentname in $attachmentnames)
	{
		[string]$attachmentpath=$PSScriptRoot + "\" + $attachmentname
		write-host "`n"
		write-host "Submitting following file:" $attachmentpath -foregroundcolor green
		if ($confirm -eq $true) 
		{
			$submit=read-host "Submit this attachment? [YES]"
		}
		else
		{
			$submit="yes"	
		}
		# Base64 encoding of the .eml file content. Reading the content of the file into a byte array.
		$EncodedContent = [Convert]::ToBase64String([IO.File]::ReadAllBytes($attachmentpath))
		
		if ($submit -eq "yes" -or $submit -eq "y")
		{
			$body = @"
			{
		"@odata.type": "#microsoft.graph.emailFileAssessmentRequest",
		"recipientEmail": "$recipient",
		"category": "$category",
		"expectedAssessment": "block",
		"contentData": "$EncodedContent"
}
"@
			$accessToken=Get-AccessToken
			$Headers= @{"Content-Type" = "application/json" ; "Authorization" = "Bearer " + $accessToken}
		
			$error.clear()
			$ErrorOccured = $false
			$ThreatRequestJSON=Invoke-WebRequest -Uri $GraphURL -Headers $headers -Body $body -Method POST -ContentType 'application/json; charset=utf-8' -ErrorAction continue -ErrorVariable ProcessError
			if($ProcessError)
			{
				write-host "Error submitting Email: " -foregroundcolor red
				write-host $ProcessError -foregroundcolor red
				$ErrorOccured=$true
				Continue
			}
			if (!$ErrorOccured)
			{
				$ThreatRequest=$ThreatRequestJSON.content |convertfrom-JSON
				write-host "Email submitted. Submission ID: " $ThreatRequest.id -foregroundcolor green
				write-host "`n"
				$ThreatRequestIDs+=$ThreatRequest.id
				$ThreatRequest.id | out-file $fileSubmissionIDs -append
				$anythingsubmitted=$true
			}
			$submit="no"
		}
	}
	
	if ($anythingsubmitted)
	{
		write-host "All request IDs:"
		$ThreatRequestIDs
		if ($confirm -eq $true) 
		{
			$checkSubmission=Read-Host "Do you want to check Submission status? [Yes\No]"
		}
		else
		{
			$checkSubmission="no"
		}
		if ($checkSubmission -eq "Yes")
		{
			Check-Submission -ThreatRequestIDs $ThreatRequestIDs
		}
	}
	
}	

Function Check-Submission {
Param
(
	$ThreatRequestIDs
)

	[int]$time=Read-Host "Submission completed. How long do you want to wait in minutes until status check? [recommended 5]"
	$delay=$time*60
	write-host "Waiting for $time minutes"
	start-sleep -seconds $delay
	foreach ($ThreatRequestID in $ThreatRequestIDs)
	{
		write-host "Checking status of submission ID: " $ThreatRequestID
		$SubmissionResultURL="$GraphURL/$ThreatRequestID"+'?$expand=results'
		
		$Headers= @{"Content-Type" = "application/json" ; "Authorization" = "Bearer " + $accessToken}
		$ThreatRequestResult=Invoke-WebRequest -Uri $SubmissionResultURL -Headers $headers -Method GET -ContentType 'application/json; charset=utf-8'
		
		$content=$ThreatRequestResult.content | convertFrom-JSON
		write-host "Current status of the submission:" -foregroundcolor green
		$content | fl id,createdDateTime,ContentType,expectedAssessment,category,status,recipientEmail,destinationRoutingReason,results
		if ($content.status -eq "completed")
		{
			write-host "Submission ID: " $content.id "submission results:" $content.results.message -foregroundcolor green
			write-host "Result of: "$content.results[0].resultType -foregroundcolor green
			$content.results[0] | select message,createdDateTime
			write-host "Result of: "$content.results[1].resultType -foregroundcolor green
			$content.results[1] | select message,createdDateTime
		}
		else
		{
			write-host "Submission ID: " $content.id " sScan still in progress" -foregroundcolor yellow
			[int]$continue=read-host 'Check again the scan result? If yes how long do you want to wait? [Press "0" for No or enter other digit for duration in minutes]'
			if ($continue -ne "0")
			{
				$delay2=$continue*60
				write-host "Waiting for $continue minutes"
				start-sleep $delay2
				$ThreatRequestResult=Invoke-WebRequest -Uri $SubmissionResultURL -Headers $headers -Method GET -ContentType 'application/json; charset=utf-8'
				$content=$ThreatRequestResult.content | convertFrom-JSON
				if ($content.status -eq "completed")
				{
					write-host "Submission ID: " $content.id "submission results:" $content.results.message -foregroundcolor green
					write-host "Result of: "$content.results[0].resultType -foregroundcolor green
					$content.results[0] | select message,createdDateTime
					write-host "Result of: "$content.results[1].resultType -foregroundcolor green
					$content.results[1] | select message,createdDateTime
				}
				else
				{
					write-host "Submission ID: " $content.id " Scan still in progress" -foregroundcolor yellow
				}
			}
			else
			{
				write-host "Submission IDs are in $fileSubmissionIDs"
			}
		}
	}
}

# Function calls:

$FindEmailResult=Find-Email -InternetMessageID $InternetMessageID -mailbox $mailbox -attachment $attachment

if ($attachment)
{
	Submit-Attachment -recipient $FindEmailResult.recipient -attachmentnames $FindEmailResult.path -category $category
}
else
{
	Submit-Email -MessageURI $FindEmailResult.MessageURI -recipient $FindEmailResult.recipient -category $category
}
