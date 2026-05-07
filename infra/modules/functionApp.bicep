@description('Azure region for all resources')
param location string

@description('Base name used for all resources (e.g. haloxteams)')
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

var deploymentContainerName = 'deployments'

// --- Storage Account (required by Azure Functions, also hosts deployment artefacts) ---
var storageAccountName = replace(toLower('st${appName}'), '-', '')
var truncatedStorageName = length(storageAccountName) > 24 ? substring(storageAccountName, 0, 24) : storageAccountName

resource storageAccount 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: truncatedStorageName
  location: location
  sku: {
    name: 'Standard_LRS'
  }
  kind: 'StorageV2'
  properties: {
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
  }
}

resource deploymentContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-05-01' = {
  name: '${storageAccount.name}/default/${deploymentContainerName}'
}

// --- App Service Plan (Flex Consumption / FC1) ---
resource appServicePlan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: '${appName}-plan'
  location: location
  kind: 'functionapp'
  sku: {
    tier: 'FlexConsumption'
    name: 'FC1'
  }
  properties: {
    reserved: true
  }
}

// --- Function App (Flex Consumption, Linux) ---
resource functionApp 'Microsoft.Web/sites@2023-12-01' = {
  name: appName
  location: location
  kind: 'functionapp,linux'
  dependsOn: [
    deploymentContainer
  ]
  properties: {
    serverFarmId: appServicePlan.id
    httpsOnly: true
    siteConfig: {
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
    functionAppConfig: {
      deployment: {
        storage: {
          type: 'blobContainer'
          value: 'https://${storageAccount.name}.blob.${environment().suffixes.storage}/${deploymentContainerName}'
          authentication: {
            type: 'StorageAccountConnectionString'
            storageAccountConnectionStringName: 'AzureWebJobsStorage'
          }
        }
      }
      scaleAndConcurrency: {
        maximumInstanceCount: 40
        instanceMemoryMB: 2048
      }
      runtime: {
        name: 'node'
        version: '24'
      }
    }
  }
}

// --- Outputs ---
output functionAppHostname string = functionApp.properties.defaultHostName
output messagingEndpoint string = 'https://${functionApp.properties.defaultHostName}/api/messages'
output webhookEndpoint string = 'https://${functionApp.properties.defaultHostName}/api/notify'
