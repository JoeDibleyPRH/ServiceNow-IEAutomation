<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.152
	 Created on:   	22/06/2018 11:07
	 Created by:   	Joe Dibley
	 Organization: 	Penguin Random House 
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		Script to log requests in Service Now
#>

param
(
	[parameter(Mandatory = $true)]
	[String]$ShortDescription,
	[parameter(Mandatory = $true)]
	[String]$LongDescription,
	[parameter(Mandatory = $true)]
	[Switch]$ReturnReqNumber,
	[parameter(Mandatory = $false)]
	[Switch]$IEVisible,
	[parameter(Mandatory = $true)]
	[String]$Username,
	[parameter(Mandatory = $false)]
	[String]$PriorityText = "3 - Moderate",
	[parameter(Mandatory = $true)]
	[String]$TrackingFile,
	[parameter(Mandatory = $true)]
	[String]$BusinessServiceID,
	[parameter(Mandatory = $true)]
	[String]$CategoryText = "AdminMaintenance",
	[parameter(Mandatory = $true)]
	[String]$SMTPServer,
	[parameter(Mandatory = $true)]
	[String]$AlertToEmailAddress,
	[parameter(Mandatory = $true)]
	[String]$AlertFromEmailAddress,
	[parameter(Mandatory = $true)]
	[String]$RequestURL,
	[parameter(Mandatory = $true)]
	[String]$GenericAccountName,
	[parameter(Mandatory = $true)]
	[String]$GenericAccountID
)



$Username = ""
$PasswordFile = "$PSScriptRoot\Password.txt"
#First need to encrypt password with powershell by doing the following
# "PASSWORD" | ConvertTo-SecureString -asplaintext -force | ConvertFrom-SecureString | Out-File <File Location>
#This must be done on the computer that the script will be configured and under the service account it will run as.
$Password = Get-Content $PasswordFile | Convertto-securestring
$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))

#Setting up error emails
function Send-ErrorEmail
{
	param
	(
		[parameter(Mandatory = $true)]
		[String]
		$body
	)
	
	$Subject = "Error: ServiceNow_RequestLogger"
	$s = New-Object System.Security.SecureString
	$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "NT AUTHORITY\ANONYMOUS LOGON", $S
	
	$messageParameters = @{
		Subject	     = $Subject
		Body		 = $body
		from		 = $AlertFromEmailAddress
		To		     = $AlertToEmailAddress
		SmtpServer   = $SMTPServer
		Credential   = $creds
	}
	Send-MailMessage @messageParameters -BodyAsHtml
	
}


$i = 0
do
{
	try
	{
		$ie = new-object -ComObject "InternetExplorer.Application" -ErrorVariable errorvar -ErrorAction Stop
		#If successful then clear out ErrorVar so does not trigger the error email function
		Clear-Variable Errorvar	
	}
	catch { Write-Host "IE Failed to open. Trying again"; Start-Sleep 5 }
	$i = $i + 1
}
until ($ie -or $i -gt 10)

if ($errorvar)
{
	Send-ErrorEmail -body "IE Failed to Open with error: $errorvar"	
}

$requestUri = $RequestURL
$userIdFragment = "user_name";
$passwordIdFragment = "user_password";
$buttonIdFragment = "sysverb_login";
if ($IEVisible)
{
	$ie.visible = $true
}
else
{
	$ie.visible = $false
}
$ie.silent = $true
$ie.navigate($requestUri)
Write-Host "Navigating to: $RequestURI"
while ($ie.Busy) { Start-Sleep -Milliseconds 100 }

$i = 0
do
{
	$ErrorActionPreference = 'SilentlyContinue'
	$doc = $ie.Document
	Start-Sleep 1
	$doc1 = $doc.iHTMLDocument3_getElementById("gsft_main")
	Start-Sleep 1
	#get inner iframe document
	$doc2 = $doc1.ContentDocument
	Write-Host "Looping to get iframe"
	Start-Sleep 1
	$i = $i + 1
}
until ($doc2 -or $i -gt 150)
$ErrorActionPreference = 'Stop'

if ($i -gt 150)
{
	Send-ErrorEmail -body "Failed to get iframe in IE"
}


Write-Host "Logging In"
$i = 0
do
{
	$ErrorActionPreference = 'SilentlyContinue'
	#Get and set username and user password options
	$UsernameField = $doc2.iHTMLDocument3_GetElementById("user_name")
	$PasswordField = $doc2.iHTMLDocument3_GetElementById("user_password")
	$LoginButton = $doc2.iHTMLDocument3_GetElementById("sysverb_login")
	
	#Prepopulate password value!
	$UsernameField.Value = $Username
	$PasswordField.Value = $Password
	Start-Sleep 1
	$i = $i + 1
}
until (($UsernameField.Value -eq $Username -and $PasswordField.Value -eq $Password) -or $ie.document.title -like "*Dashboard*" -or $i -gt 150)

if ($i -gt 150)
{
	Send-ErrorEmail -body "Failed to Login or detect auto login"
}

$ErrorActionPreference = 'Stop'

if ($ie.document.title -notlike "*Dashboard*")
{
	$LoginButton.click()
}

while ($ie.Busy) { Start-Sleep 1; Write-Host "IE is busy: $($ie.busy)" }

while ($ie.document.title -notlike "*Dashboard*") { Start-Sleep -Milliseconds 100; Write-Host "Waiting for tab to be like *Dashboard*: $($ie.document.title)" }

Write-Host "Login Complete"


Write-Host "Getting Document"
$doc = $ie.Document
Write-Host "Getting Favorites_tab and clicking"
$button = $doc.iHTMLDocument3_GetElementById("favorites_tab")
Write-Host "Clicking Favourites tab"
$button.click()

while ($ie.Busy) { Start-Sleep -Seconds 1; Write-host "IE is Busy: $($ie.busy)" }
$i = 0
do
{
	$ErrorActionPreference = 'SilentlyContinue'
	$doc = $ie.Document
	$doc1 = $doc.iHTMLDocument3_GetElementById("gsft_nav")
	Start-Sleep 2
	$elements = $doc1.GetElementsByClassName("sn-widget-list-title nav-favorite-title ng-binding nav-favorite-TABLE")
	Start-Sleep 2
	$elements[0].Click()
	Write-Host "Attempting to get New Request favourite"
	$i = $i + 1
}
until ($ie.document.title -like "Req*" -or $i -gt 150)

$ErrorActionPreference = 'Stop'

if ($i -gt 150)
{
	Send-ErrorEmail -body "Failed to click on New Request Favourite"
}

while ($ie.Busy) { Start-Sleep -Milliseconds 100 }

#Setting Assignment Group
Write-Host "Getting iframe"
$i = 0
do
{
	$ErrorActionPreference = 'SilentlyContinue'
	$doc = $ie.Document
	$doc1 = $doc.iHTMLDocument3_GetElementById("gsft_main")
	$doc2 = $doc1.ContentDocument
	$i = $i + 1
}
until ($doc2 -or $i -gt 150)
$ErrorActionPreference = 'Stop'

if ($i -gt 150)
{
	Send-ErrorEmail -body "Failed to retrieve the gsft_main iframe for new request"
}

Write-Host "Filling in request"
#Setting Value for Contact Type to "Self-Service"
$doc3 = $doc2.iHTMLDocument3_GetElementById("sc_request.contact_type")
$doc3.Options.SelectedIndex = 3

#Setting requested by to generic account
$doc3 = $doc2.iHTMLDocument3_GetElementById("sc_request.u_requested_by")
$doc3.Value = $GenericAccountID
$doc3.DefaultValue = $GenericAccountID

#Setting requested for to generic account
$doc3 = $doc2.iHTMLDocument3_GetElementById("sc_request.u_requested_for")
$doc3.Value = $GenericAccountName

#Setting Short description
$doc3 = $doc2.iHTMLDocument3_GetElementById("sc_request.short_description")
$doc3.Value = $ShortDescription

#Setting actual description
$doc3 = $doc2.iHTMLDocument3_GetElementById("sc_request.description")
$doc3.Value = $LongDescription

#Setting request priority
$doc3 = $doc2.iHTMLDocument3_GetElementById("sc_request.priority")
$doc3.Options.SelectedIndex = ($doc3.Options | where { $_.Value -eq $PriorityText } | select -expand Index)


#Setting Technical Aplication / Service to "Servers Virtual - Windows"
$doc3 = $doc2.iHTMLDocument3_GetElementById("sc_request.business_service")
$doc3.Value = $BusinessServiceID
$doc3.DefaultValue = $BusinessServiceID

$doc3 = $doc2.iHTMLDocument3_GetElementById("sc_request.u_category1")
$doc3.Options.SelectedIndex = ($doc3.Options | where { $_.Value -eq $CategoryText } | select -expand Index)

#Getting Request Number
$doc3 = $doc2.iHTMLDocument3_GetElementById("sys_readonly.sc_request.number")
$RequestNumber = $doc3.Value

$doc3 = $doc2.iHTMLDocument3_GetElementById("sysverb_update")
$doc3.Click()

if ($ReturnReqNumber)
{ Write-Output  $RequestNumber}

