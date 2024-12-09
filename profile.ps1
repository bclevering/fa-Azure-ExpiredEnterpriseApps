# Azure Functions profile.ps1
#
# This profile.ps1 will get executed every "cold start" of your Function App.
# "cold start" occurs when:
#
# * A Function App starts up for the very first time
# * A Function App starts up after being de-allocated due to inactivity
#
# You can define helper functions, run commands, or specify environment variables
# NOTE: any variables defined that are not environment variables will get reset after the first execution

# Authenticate with Azure PowerShell using MSI.
# Remove this if you are not planning on using MSI or Azure PowerShell.
if ($env:MSI_SECRET) {
  try {
    Write-Verbose "Authenticating with Azure PowerShell using Managed Identity."
    Disable-AzContextAutosave -Scope Process | Out-Null
    Connect-AzAccount -Identity | Out-Null
  } catch {
    Write-Error "Failed to authenticate with Azure PowerShell using Managed Identity."
    throw $_
  }
}

# Authenticate with Microsoft Graph using MSI.
if ($env:MSI_SECRET) {
  try {
    Write-Verbose "Authenticating with Microsoft Graph using Managed Identity."
    Connect-MgGraph -Identity -NoWelcome
  } catch {
    Write-Error "Failed to authenticate with Microsoft Graph using Managed Identity."
    throw $_
  }
}

# Uncomment the next line to enable legacy AzureRm alias in Azure PowerShell.
# Enable-AzureRmAlias

# You can also define functions or aliases that can be referenced in any of your PowerShell functions.

class ExpiredAppCredentials {
  [String] $SecretName
  [DateTime] $ExpirationTime
  # hidden [TimeSpan] $Remaning
  [Int] $RemainingDays
  [Bool] $Expired
}

class ExpiredAppInformation {
  [String] $ApplicationName
  [String] $ApplicationID
  [String] $ApplicationObjectId
  [String] $OwnerId
  [String] $OwnerUsername
  [ExpiredAppCredentials[]] $ExpiredSecrets
}