# This file enables modules to be automatically managed by the Functions service.
# See https://aka.ms/functionsmanageddependency for additional information.
#
@{
    # For latest supported version, go to 'https://www.powershellgallery.com/packages/Az'. 
    # To use the Az module in your function app, please uncomment the line below.
    # 'Az' = '12.*'
    # Microsoft Graph Modules
    'Microsoft.Graph.Users'            = '2.*'
    'Microsoft.Graph.Users.Actions'    = '2.*'
    'Microsoft.Graph.Groups'           = '2.*'
    'Microsoft.Graph.DeviceManagement' = '2.*'
    'Microsoft.Graph.Sites'            = '2.*'
    # 'Microsoft.Graph.Teams'            = '2.*'
    
    # Exchange Online Management Module
    'ExchangeOnlineManagement'         = '3.*'
}
