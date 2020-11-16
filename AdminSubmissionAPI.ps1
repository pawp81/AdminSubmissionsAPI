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
#>

#Parameters
param(
        [Parameter(Mandatory = $true)]
        [string]$Username,
		[string]$url,
		[string]$filepath,
		[boolean]$checkURL,
		[string]$emailaddress,
		[string]$MessageID
		
    )
	
# configure below variables according to: https://github.com/pawp81/AdminSubmissionsAPI	
$clientId = ""
$tenantId = ""

# List of static variables.
$resourceURI = "https://graph.microsoft.com"
$GraphUrl="https://graph.microsoft.com/v1.0/informationProtection/threatAssessmentRequests"
$day=(get-date).day
$month=(get-date).month
$year=(get-date).year
$random=get-random -Maximum 10000
$fileSubmissionIDs="\SubmissionIDs-$domain-$day-$month-$year_$random.txt" #specify path were submission IDs should be exported. For example: c:\Submissions\SubmissionIDs-$domain-$day-$month-$year.txt

#Function to obtain access token and save it on disk as accessToken.txt.
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

	
function Submit-URL {
# Function to submit URL from the file or single URL to the Microsoft.
	Param
	(	
		$URLArray 
	)	
	$ThreatRequestIDs=@()
	
	foreach ($URLarray_ID in $URLarray)
	{
	
		write-host "Submitting $URLarray_ID"  -foregroundcolor green
		if ($checkURL)
		{
			start-process microsoft-edge:$URLarray_ID # launching Edge
	
		}
		$body = @"
		{
	"@odata.type": "#microsoft.graph.urlAssessmentRequest",
	"url": "$URLarray_ID",
	"category": "phishing",
	"expectedAssessment": "block"
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
				write-host "Error submitting URL: " -foregroundcolor red
				write-host $_ -foregroundcolor red
				$ErrorOccured=$true
			}
			if (!$ErrorOccured)
			{
				$ThreatRequest=$ThreatRequestJSON.content |convertfrom-JSON
				write-host "URL submitted. Submission ID: " $ThreatRequest.id -foregroundcolor green
				write-host "`n"
				$ThreatRequestIDs+=$ThreatRequest.id
				$ThreatRequest.id | out-file $fileSubmissionIDs -append
				$anythingsubmitted=$true
			}
	}
	write-host "All request IDs:"
	$ThreatRequestIDs
	#Checking status of the URL submission
	if ($anythingsubmitted)
	{
		$checkSubmission=Read-Host "Do you want to check Submission status? [Yes\No]"
		if ($checkSubmission -eq "Yes")
		{
			[int]$time=Read-Host "Submission completed. How long do you want to wait in minutes until status check? [recommended 5]"
			$delay=$time*60
			write-host "waiting for $time minutes"
			start-sleep -seconds $delay
			foreach ($ThreatRequestID in $ThreatRequestIDs)
			{
				write-host "Checking status of submission ID: " $ThreatRequestID
				$SubmissionResultURL="$GraphURL/$ThreatRequestID"+'?$expand=results'
				
				$Headers= @{"Content-Type" = "application/json" ; "Authorization" = "Bearer " + $accessToken}
				$ThreatRequestResult=Invoke-WebRequest -Uri $SubmissionResultURL -Headers $headers -Method GET -ContentType 'application/json; charset=utf-8'
				
				$content=$ThreatRequestResult.content | convertFrom-JSON
				if ($content.results.message)
				{
					write-host $content.url " submission results:" $content.results.message -foregroundcolor green
				}
				else
				{
					write-host $content.url " scan still in progress" -foregroundcolor yellow
					[int]$continue=read-host 'Check again for scan result? If yes how long do you want to wait? [Press "0" for No or enter other digit for duration in minutes]'
					if ($continue -ne "0")
					{
						$delay2=$continue*60
						write-host "waiting for $continue minutes"
						start-sleep $delay2
						$ThreatRequestResult=Invoke-WebRequest -Uri $SubmissionResultURL -Headers $headers -Method GET -ContentType 'application/json; charset=utf-8'
						$content=$ThreatRequestResult.content | convertFrom-JSON
						if ($content.results.message)
						{
							write-host $content.url " submission results:" $content.results.message -foregroundcolor green
						}
						else
						{
							write-host $content.url " scan still in progress" -foregroundcolor yellow
						}
					}
				}
			}
		}
		else
		{
			write-host "Submission IDs are in $fileSubmissionIDs"
		}
	}
}

if ($filepath)
{
	$URLarray=get-content $filepath
	Submit-URL -UrlArray $URLarray
}
else
{
	if ($url)
	{
		Submit-URL -UrlArray $url
	}
	else
	{
		write-host "No URL was provided for Submission. Quitting"
	}
}
