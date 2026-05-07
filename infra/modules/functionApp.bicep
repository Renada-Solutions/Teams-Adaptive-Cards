@description('Azure region for all resources')
param location string

@description('Base name used for all resources (e.g. halopsa-ooh-bot)')
param appName string

@description('Application (Client) ID from Azure AD App Registration')
param microsoftAppId string

@secure()
@description('Client secret from Azure AD App Registration')
param microsoftAppPassword string

@description('Directory (Tenant) ID from Azure AD')
param microsoftAppTenantId string

@secure()
@description('Shared secret for authenticating HaloPSA webhook calls (Bearer token)')
param notifySecret string

@description('GitHub repository URL containing the bot code')
param repoUrl string

@description('GitHub branch to deploy from')
param repoBranch string

// --- Storage Account (required by Azure Functions) ---
var storageAccountName = replace(toLower('st${appName}'), '-', '')
var truncatedStorageName = length(storageAccountName) > 24 ? substring(storageAccountName, 0, 24) : storageAccountName

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: truncatedStorageName
  location: location
  sku: {
    name: 'Standard_LRS'
  }
  kind: 'StorageV2'
}

// --- App Service Plan (Consumption / Y1) ---
resource appServicePlan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: '${appName}-plan'
  location: location
  sku: {
    name: 'Y1'
    tier: 'Dynamic'
  }
  properties: {
    reserved: false
  }
}

// --- Function App ---
resource functionApp 'Microsoft.Web/sites@2023-12-01' = {
  name: appName
  location: location
  kind: 'functionapp'
  properties: {
    serverFarmId: appServicePlan.id
    httpsOnly: true
    siteConfig: {
      nodeVersion: '~24'
      use32BitWorkerProcess: false
      ftpsState: 'Disabled'
      http20Enabled: true
      minTlsVersion: '1.3'
      scmMinTlsVersion: '1.3'
      appSettings: [
        {
          name: 'AzureWebJobsStorage'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storageAccount.listKeys().keys[0].value}'
        }
        {
          name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storageAccount.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storageAccount.listKeys().keys[0].value}'
        }
        {
          name: 'WEBSITE_CONTENTSHARE'
          value: toLower(appName)
        }
        {
          name: 'FUNCTIONS_EXTENSION_VERSION'
          value: '~4'
        }
        {
          name: 'FUNCTIONS_WORKER_RUNTIME'
          value: 'node'
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~24'
        }
        {
          name: 'PROJECT'
          value: 'bot'
        }
        {
          name: 'MicrosoftAppId'
          value: microsoftAppId
        }
        {
          name: 'MicrosoftAppPassword'
          value: microsoftAppPassword
        }
        {
          name: 'MicrosoftAppTenantId'
          value: microsoftAppTenantId
        }
        {
          name: 'NOTIFY_SECRET'
          value: notifySecret
        }
      ]
    }
  }
}

// --- Source Control Deployment from GitHub ---
resource sourceControl 'Microsoft.Web/sites/sourcecontrols@2023-12-01' = {
  parent: functionApp
  name: 'web'
  properties: {
    repoUrl: repoUrl
    branch: repoBranch
    isManualIntegration: true
  }
}

// --- Outputs ---
output functionAppHostname string = functionApp.properties.defaultHostName
output messagingEndpoint string = 'https://${functionApp.properties.defaultHostName}/api/messages'
output webhookEndpoint string = 'https://${functionApp.properties.defaultHostName}/api/notify'
