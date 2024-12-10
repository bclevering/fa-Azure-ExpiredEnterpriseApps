# Input bindings are passed in via param block.
param($Timer)

$ErrorActionPreference = "Stop"

# The 'IsPastDue' property is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
  Write-Host "PowerShell timer is running late!"
}

# Write an information log with the current time.
$currentUTCtime = (Get-Date).ToUniversalTime()
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"

## Check existance of required environment variables
if (-Not (Test-Path -Path ENV:API_FUNCTION_KEY)) {
  Write-Warning "API_FUNCTION_KEY environment variable is not set. Calling the backend API may fail."
  exit 1
}
if (-Not (Test-Path -Path ENV:WEBSITE_HOSTNAME)) {
  Write-Error "WEBSITE_HOSTNAME environment variable is not set."
  exit 1
}

try {

  $apiFunctionKey = $env:API_FUNCTION_KEY
  Write-Host "Calling API at $apiEndpointUrl with Function Key. $($apiFunctionKey)"

  $apiEndpointUrl = "https://$($env:WEBSITE_HOSTNAME)/api/GetExpiredSecrets?code=$($apiFunctionKey)"
  $expiredSecrets = Invoke-RestMethod -Method Get -Uri $apiEndpointUrl -OperationTimeoutSeconds 60 -ConnectionTimeoutSeconds 65

  $expiredSecrets | Select-Object -Property ApplicationName, OwnerUsername -ExpandProperty ExpiredSecrets
} catch {
  Write-Error "Failed to get expired secrets from API."
  throw $_
}

$htmlTable = $expiredSecrets |
Select-Object -ExpandProperty ExpiredSecrets |
ConvertTo-Html -Fragment


$message = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>HTML TABLE</title>
</head>
<body>
<h1>Expired Secrets:</h1>
$htmlTable
</body>
</html>
"@

$MailFrom = "noreply@clevering.eu"
$MailTo = "bert@clevering.eu"
$msgBody = $message

$Message = @{
   Subject = "Weekly report for expiring Enterprise apps secrets"
   Body = @{
      ContentType = "HTML"
      Content = $msgBody
      }
   ToRecipients = @(
      @{
         EmailAddress = @{
         Address = $MailTo
         }
       }
    )
}

Send-MgUserMail -UserId $MailFrom -Message $Message

$message