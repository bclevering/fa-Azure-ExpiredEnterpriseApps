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
  $apiEndpointUrl = "https://$($env:WEBSITE_HOSTNAME)/api/GetExpiredSecrets"
  $apiFunctionKey = $env:API_FUNCTION_KEY

  Write-Host "Calling API at $apiEndpointUrl with Function Key."
  $expiredSecrets = Invoke-RestMethod -Method Get `
    -Headers @{
    "x-functions-key" = $apiFunctionKey
  } `
    -Uri $apiEndpointUrl

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

$message