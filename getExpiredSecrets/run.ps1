# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

$ErrorActionPreference = "Stop"

function Get-ApplicationOwner {
  param (
    [Parameter(
      Position = 0,
      Mandatory = $true
    )]
    [string]$ApplicationObjectId
  )

  Write-Verbose "Retrieving Owner information of application with object id: $ApplicationObjectId"

  $Owner = Get-MgApplicationOwner -ApplicationId $ApplicationObjectId
  $OwnerID = $Owner.Id
  $Username = $Owner.AdditionalProperties.displayName

  if ($null -eq $Owner.AdditionalProperties.displayName) {
    $Username = $Owner.AdditionalProperties.userPrincipalName
  }
  if ($null -eq $Owner.AdditionalProperties.userPrincipalName) {
    $Username = '<<No Owner>>'
  }

  return [PSCustomObject]@{
    Id       = $OwnerID
    Username = $Username
  }
}


# Print some information about the request
Write-Host "PowerShell HTTP trigger function processed a request."

$requestMethod = $Request.Method
$remoteAddress = $Request.Headers["X-Forwarded-For"]

Write-Host "Received $requestMethod request from $remoteAddress. User Agent: $($Request.Headers["User-Agent"])"

# Process the input data
# ...

$Now = Get-Date

try {
  $DueDays = $Request.Body.days

  if (-not $DueDays) {
    Write-Host "No days parameter provided, using default value of 30 days."
    $DueDays = 30
  }
} catch {
  Write-Error "Failed to retrieve days parameter from request body."
  Push-OutputBinding -Name response -Value ([HttpResponseContext]@{
      StatusCode = [System.Net.HttpStatusCode]::BadRequest
      Body       = "Failed to retrieve days parameter from request body."
    }) -Clobber
  throw $_
}

Write-Host "Retrieving all Azure AD applications with secrets that are due to expire in $DueDays days or less."

## Retrieve all Azure AD applications and filter them by secrets to be expired
try {
  $AppsToExpire = Get-MgApplication -All -ErrorAction Stop | ForEach-Object {

    Write-Host "Processing application `"$($_.DisplayName)`"."

    $AppName = $PSItem.DisplayName
    $AppId = $PSItem.AppId
    $AppObjectId = $PSItem.Id

    # ToDo: Also check for certificates that are due to expire (KeyCredentials)
    $AppCredentials = Get-MgApplication -ApplicationId $AppObjectId | Select-Object PasswordCredentials
    $Secrets = $AppCredentials.PasswordCredentials

<#     $ExpiredSecrets = New-Object -TypeName System.Collections.Generic.List[ExpiredAppCredentials]

    foreach ($secret in $Secrets) {
      $SecretName = $secret.DisplayName
      $ExpirationTime = $secret.EndDateTime
      $Remaining = $ExpirationTime - $Now

      if ($Remaining.Days -le $DueDays) {
        Write-Host "Secret `"$($secret.DisplayName)`" is due to expire in $($Remaining.Days) days."

        $ExpiredSecrets.Add([ExpiredAppCredentials]@{
            SecretName     = $SecretName
            ExpirationTime = $ExpirationTime
            # Remaning       = $Remaining
            RemainingDays  = $Remaining.Days
            Expired        = $Remaining.TotalSeconds -le 0
          })
      }
    } #>

    # Return if the application has no secrets to expire
    if ($ExpiredSecrets.Count -eq 0 ) {
      Write-Host "Application `"$AppName`" has no secrets to expire."
      return
    }

<#     $Owner = Get-ApplicationOwner -ApplicationObjectId $AppObjectId

    return [ExpiredAppInformation]@{
      ApplicationName     = $AppName
      ApplicationID       = $AppId
      ApplicationObjectId = $AppObjectId
      OwnerId             = $Owner.Id
      OwnerUsername       = $Owner.Username
      ExpiredSecrets      = $ExpiredSecrets
    } #>
  }
} catch {
  Write-Error "Failed to retrieve Azure AD applications."
  Push-OutputBinding -Name response -Value ([HttpResponseContext]@{
      StatusCode = [System.Net.HttpStatusCode]::InternalServerError
      Body       = "Failed to retrieve Azure AD applications."
    }) -Clobber
  throw $_
}

Push-OutputBinding -Name response -Value ([HttpResponseContext]@{
    StatusCode = [System.Net.HttpStatusCode]::OK
    Body       = "DONE"
    #$appsToExpire | ConvertTo-Json -Depth 10
  }) -Clobber