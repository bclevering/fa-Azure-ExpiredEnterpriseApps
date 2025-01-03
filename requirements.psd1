# This file enables modules to be automatically managed by the Functions service.
# See https://aka.ms/functionsmanageddependency for additional information.
#
@{
  # For latest supported version, go to 'https://www.powershellgallery.com/packages/Az'.
  # To use the Az module in your function app, please uncomment the line below.
  'Az.Accounts'                    = '2.*'
  'Az.Storage'                     = '5.*'
  'Microsoft.Graph.Authentication' = '2.*'
  'Microsoft.Graph.Applications'   = '2.*'
  'Microsoft.Graph.Users.Actions'  = '2.*'
  'Microsoft.Graph.Groups'         = '2.*'
}