  
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
 Submit attachments of the emails received by soc@contoso.com mailbox in last 24h. Submit emails as phishing. Use admin@contoso.com credentials for authentication
.\Email_AdminSubmissionAPI.ps1 -Username admin@contoso.com -agoHours 24 -mailbox soc@contoso.com -category phishing -attachment $true

 Submit emails received by John Doe mailbox in last 24h as phishing.
.\Email_AdminSubmissionAPI.ps1 -Username admin@contoso.com  -agoHours 24 -mailbox john.doe@contoso.com -category phishing

 Submit emails received by John Doe mailbox in last 24h as phishing.
.\Email_AdminSubmissionAPI.ps1 -Username admin@contoso.com  -agoHours 24 -mailbox john.doe@contoso.com -category phishing
#>

param(
        [Parameter(Mandatory = $true)]
        [string]$Username,
		[int]$agoHours,
		[int]$agoMinutes,
		[string]$InternetMessageID,
		[Parameter(Mandatory = $true)]
		[string]$mailbox,
		[boolean]$attachment,
		[Parameter(Mandatory = $true)]
		[string]$category, #allowed values: "phishing", "spam"
		[boolean] $confirm #no prompt for making a submission.
    )

# configure below variables according to: https://github.com/pawp81/AdminSubmissionsAPI
$clientId =  ""
$tenantId = "" 


# List of static variables.
$resourceURI = "https://graph.microsoft.com"
$GraphUrl="https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests"
$day=(get-date).day
$month=(get-date).month
$year=(get-date).year
$random=get-random -Maximum 10000
$fileSubmissionIDs="SubmissionIDs-$day-$month-$year_$random.txt"
$logname="Log-$day-$month-$year_$random.txt"

function Get-AccessToken{
	
	try
	{
		write-host "Obtaining accessToken"
		$MsalClientApp = New-MsalClientApplication -ClientId $clientId -TenantId $tenantId | Enable-MsalTokenCacheOnDisk -PassThru
		$MsalToken = $MsalClientApp | Get-MsalToken  -Scopes "$resourceURI/Mail.Read","$resourceURI/Mail.Read.Shared","$resourceURI/ThreatAssessment.ReadWrite.All","$resourceURI/User.Read"
		$accessToken=$MsalToken.AccessToken
	}
	catch
	{
	}
		
	return $accessToken
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
		$URI = "https://graph.microsoft.com/v1.0/users/$mailbox/mailFolders/inbox/messages?`$top=1000&`$select=id,toRecipients&`$filter=InternetMessageID eq '<$InternetMessageID>'"
	}
	if ($agoHours)
	{
		$startdate_temp=(get-date).AddHours(-$agohours).ToUniversalTime()
		$startdate=$startdate_temp.tostring("yyyy-MM-dd'T'HH:mm'Z'")
		$URI = "https://graph.microsoft.com/v1.0/users/$mailbox/mailFolders/inbox/messages?`$top=1000&`$select=id,toRecipients&`$filter=ReceivedDateTime ge $startdate"
	}
	if ($agoMinutes)
	{
		$startdate_temp=(get-date).AddMinutes(-$agominutes).ToUniversalTime()
		$startdate=$startdate_temp.tostring("yyyy-MM-dd'T'HH:mm'Z'")
		$URI = "https://graph.microsoft.com/v1.0/users/$mailbox/mailFolders/inbox/messages?`$top=1000&`$select=id,toRecipients&`$filter=ReceivedDateTime ge $startdate"
	}
	
	write-host "Searching for:" $URI
	"Searching for: $URI" | out-file $logname -append
	$MessageJSON=Invoke-WebRequest -Uri $URI -Headers $headers
	$Messages=$MessageJSON.content | ConvertFrom-JSON
	write-host $Messages.value
	if ($Messages.value.length -eq 0)
	{
		write-host "Message not found. Exiting" -foregroundcolor Red
		"Message not found. Exiting" | out-file $logname -append
		exit
	}
	$MessageIDs = @()
	foreach ($Message in $Messages)
	{
		#Saving all messageIDs matching search criteria to the array
		
		$MessageIDs+=$Message.value.id
		
	}
	$MessageIDs | out-file $logname -append
	
	if ($attachment -eq $true)
	{
		#Looking for attachment ID. I need to download it to the disk.
		$path = @()
		foreach ($MessageID in $MessageIDs)
		{
			$MessageAttachmentURI="https://graph.microsoft.com/v1.0/users/$mailbox/messages/$MessageID/attachments/"
			$MessageAttachmentJSON=Invoke-WebRequest -Uri $MessageAttachmentURI -Headers $headers
			$MessageAttachment= $MessageAttachmentJSON | ConvertFrom-JSON
			write-host "Message attachment ID found: " $MessageAttachment.value.id -foregroundcolor green
			$AttachmentID=$MessageAttachment.value.id
			$AttachmentName=$MessageAttachment.value.name
						
			#downloading the attachment to current folder
			$temppath= $attachmentID+".eml"
			$AttachmentFetchURL="https://graph.microsoft.com/v1.0/users/$mailbox/messages/$MessageID/attachments/$AttachmentID/`$value"
			write-host "Sending request to fetch attachment: " $AttachmentFetchURL
			"Sending request to fetch attachment: $AttachmentFetchURL " | out-file $logname -append
			Invoke-WebRequest -Uri $AttachmentFetchURL -Headers $headers -outfile $temppath
			
			$path += $attachmentID+".eml"
			 	
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
		else {
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
		if ($confirm -eq $true)
		{
			$checkSubmission=Read-Host "Do you want to check Submission status? [Yes\No]"
			if ($checkSubmission -eq "Yes")
			{
				Check-Submission -ThreatRequestIDs $ThreatRequestIDs
			}
		}
	}
}

function Submit-Attachment {
# Function to submit email attachment
Param
(	
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
		else {
			$submit="yes"
		}
		# Base64 encoding of the .eml file content. Reading the content of the file into a byte array.
		$EncodedContent = [Convert]::ToBase64String([IO.File]::ReadAllBytes($attachmentpath))
		$a=get-content $attachmentpath
		$b=($a | Select-string -pattern "To:")
		$recipient=($b.toString() | Select-String "[a-zA-Z0-9_.-]+@[a-zA-Z0-9_.-]+").Matches.value
		if ($submit -eq "yes")
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
		else {
			$checkSubmission = "No"
		}	
		if ($checkSubmission -eq "Yes")
		{
			Check-Submission -ThreatRequestIDs $ThreatRequestIDs
		}
	}
	if ($anythingsubmitted)
	{
		write-host "`n"
		write-host "Deleting temporary .eml files" -foregroundcolor green
		foreach ($attachmentname in $attachmentnames)
		{
		
			[string]$attachmentpath=$PSScriptRoot + "\" + $attachmentname
			Remove-item $attachmentpath
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
	Submit-Attachment -attachmentnames $FindEmailResult.path -category $category
}
else
{
	Submit-Email -MessageURI $FindEmailResult.MessageURI -recipient $FindEmailResult.recipient -category $category
}
