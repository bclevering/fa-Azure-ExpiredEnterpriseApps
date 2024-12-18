# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' porperty is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"


## Check existance of required environment variables
if (-Not (Test-Path -Path ENV:API_FUNCTION_KEY)) {
    Write-Warning "API_FUNCTION_KEY environment variable is not set. Calling the backend API may fail."
    exit 1
  }
  if (-Not (Test-Path -Path ENV:SEND_FROM)) {
    Write-Warning "No send from address specified. So quiting"
    exit 1
  }
  if (-Not (Test-Path -Path ENV:SEND_TO_ExpiredGroups)) {
    Write-Warning "No send to address specified. So quiting"
    exit 1
  }
  if (-Not (Test-Path -Path ENV:WEBSITE_HOSTNAME)) {
    Write-Error "WEBSITE_HOSTNAME environment variable is not set."
    exit 1
  }
  
  
  $ErrorActionPreference = "Stop"
  $DueDays = 30
  
  if (-Not (Test-Path -Path ENV:DUE_DAYS)) {
    $DueDays = $env:DUE_DAYS
  }
  
  $Now = Get-Date
  
  Write-Host "Retrieving all M365 groups that are due to expire in $DueDays days or less."
  
  ## Retrieve all Azure AD applications and filter them by secrets to be expired
  try {
    $GroupsToExpire = Get-MgGroup -All -ErrorAction $ErrorActionPreference | ForEach-Object {
      $GroupName = $PSItem.DisplayName
      Write-Host "Processing group `"$($GroupName)`"."

      $GroupExpirationTime = $PSItem.ExpirationDateTime 
      $GroupRemaining = $GroupExpirationTime - $Now
  
      $ExpiredGroups = New-Object -TypeName System.Collections.Generic.List
  
        if ($GroupRemaining.Days -le $DueDays) {
  
          $ExpiredGroups.Add(@{
              GroupName     = $GroupName
              ExpirationTime = $GroupExpirationTime
              RemainingDays  = $GroupRemaining.Days
              Expired        = $GroupRemaining.TotalSeconds -le 0
            })
        }
      }
  
      # Return if the application has no secrets to expire
      if ($ExpiredGroups.Count -eq 0 ) {
         Write-Host "No Groups has to expire."
         return
      }
  
      return $ExpiredGroups
  } catch {
    Write-Error "Failed to retrieve M365 groups."
    throw $_
  }
  
  
  
  try {
    $htmlTable = $GroupsToExpire |
    ConvertTo-Html -Fragment
  
    $mailTo =  $env:SEND_TO_ExpiredGroups
    $mailFrom = $env:SEND_FROM
  
    $mailMessage = 
@" 
    <!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN'  'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>
    <html xmlns='http://www.w3.org/1999/xhtml'>
    <head>
    <title>HTML TABLE</title>
    </head>
    <body>
    <h1>Expired and bound to expire M365 Groups:</h1>
    $htmlTable
    </body>
    </html>
"@
  
    $msgBody = $mailMessage
  
    $params = @{
      message = @{
        Subject = "Weekly report for expiring M365 Groups"
        Body = @{
            ContentType = "HTML"
            Content = $msgBody
            }
        ToRecipients = @(
            @{
              EmailAddress = @{
              Address = $mailTo
              }
            }
          )
      }
      saveToSentItems = "false"
    }
  
    Send-MgUserMail -userid $mailFrom -BodyParameter $params
  }
  catch {
    Write-Error "Failed to send e-mail."
    throw $_
  }
  
  
  $mailMessage