  
<#
	.SYNOPSIS
		Get Admin Submission status via the API
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
		
	.EXAMPLE
	Check submission status of single submission:
		.\Get-AdminSubmission.ps1 -username admin@contoso.com -SubmissionID 286ed0f3-ce7d-4cf1-5279-08d87738010a
	
	Check submission statuses if multiple submissions by importing them from .txt file. Txt file should contain single SubmissionID by per line.
		.\Get-AdminSubmission.ps1 -username admin@contoso.com -path SubmissionResult-10-10-2020_1357.txt
#>

#Parameters
param(
        [Parameter(Mandatory = $true)]
        [string]$Username,
		[string]$SubmissionID,
		[string]$path
		
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
$fileSubmissionIDs="c:\PawelSecurityReader\SubmissionResult-$day-$month-$year_$random.csv"

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

function Get-Submission{
Param
	(	
		$ThreatRequestID
	
	)
	$accessToken=Get-AccessToken	
	write-host "Checking status of submission ID: " $ThreatRequestID
	$SubmissionResultURL="$GraphURL/$ThreatRequestID"+'?$expand=results'
			
	$Headers= @{"Content-Type" = "application/json" ; "Authorization" = "Bearer " + $accessToken}
	$ThreatRequestResult=Invoke-WebRequest -Uri $SubmissionResultURL -Headers $headers -Method GET -ContentType 'application/json; charset=utf-8'
	$content=$ThreatRequestResult.content | convertFrom-JSON
	$content | select id,CreatedDateTime,contentType,expectedAssessment,category,status,requestSource,RecipientEmail,DestinationRoutingReason
	write-host "Result of: "$content.results[0].resultType -foregroundcolor green
	$content.results[0] | select message,createdDateTime
	$content.results[0] | select @{name="SubmissionID";expression={$ThreatRequestID}},@{name="ResultOf";expression={"PolicyCheck"}},message,createdDateTime | export-csv $fileSubmissionIDs -append
	if ($content.status -ne "pending")
	{	
		write-host "Result of: "$content.results[1].resultType -foregroundcolor green
		$content.results[1] | select message,createdDateTime
		$content.results[1] | select @{name="SubmissionID";expression={$ThreatRequestID}},@{name="ResultOf";expression={"Rescan"}},message,createdDateTime | export-csv $fileSubmissionIDs -append
	}
}
if ($path)
{
	$SubmissionIDs=get-content $path
	foreach ($SubmissionID in $SubmissionIDs)
	{
		Get-Submission -ThreatRequestID $SubmissionID
	}
}
else
{
	Get-Submission -ThreatRequestID $SubmissionID
}