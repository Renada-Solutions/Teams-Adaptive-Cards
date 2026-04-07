@description('Display name for the Azure Bot resource')
param botName string

@description('Application (Client) ID from Azure AD App Registration')
param microsoftAppId string

@description('Directory (Tenant) ID from Azure AD')
param microsoftAppTenantId string

@description('Bot Framework messaging endpoint URL')
param messagingEndpoint string

// --- Azure Bot (F0 Free Tier) ---
resource bot 'Microsoft.BotService/botServices@2022-09-15' = {
  name: botName
  location: 'global'
  sku: {
    name: 'F0'
  }
  kind: 'azurebot'
  properties: {
    displayName: botName
    endpoint: messagingEndpoint
    msaAppId: microsoftAppId
    msaAppTenantId: microsoftAppTenantId
    msaAppType: 'SingleTenant'
  }
}

// --- Microsoft Teams Channel ---
resource teamsChannel 'Microsoft.BotService/botServices/channels@2022-09-15' = {
  parent: bot
  name: 'MsTeamsChannel'
  location: 'global'
  properties: {
    channelName: 'MsTeamsChannel'
    properties: {
      isEnabled: true
    }
  }
}
