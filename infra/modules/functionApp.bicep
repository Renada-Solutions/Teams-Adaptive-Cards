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

@description('URL to the bot.zip package — defaults to the latest GitHub release')
param packageUrl string = 'https://github.com/Renada-Solutions/Teams-Adaptive-Cards/releases/latest/download/bot.zip'

var deploymentContainerName = 'deployments'
var deployIdentityName = '${appName}-deploy-id'
var deployScriptName = '${appName}-deploy-code'
var websiteContributorRoleId = 'de139f84-1756-47ae-9be6-808fbbe84772'

// --- Storage Account (function runtime + deployment artefact container) ---
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

// --- One-time identity + role + script that pulls the bot package into this Function App ---
resource deployIdentity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  name: deployIdentityName
  location: location
}

resource deployRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(resourceGroup().id, deployIdentityName, 'WebsiteContributor')
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', websiteContributorRoleId)
    principalId: deployIdentity.properties.principalId
    principalType: 'ServicePrincipal'
  }
}

resource deployScript 'Microsoft.Resources/deploymentScripts@2023-08-01' = {
  name: deployScriptName
  location: location
  kind: 'AzureCLI'
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${deployIdentity.id}': {}
    }
  }
  dependsOn: [
    functionApp
    deployRoleAssignment
  ]
  properties: {
    azCliVersion: '2.61.0'
    timeout: 'PT15M'
    retentionInterval: 'P1D'
    cleanupPreference: 'OnSuccess'
    environmentVariables: [
      { name: 'PACKAGE_URL', value: packageUrl }
      { name: 'RG', value: resourceGroup().name }
      { name: 'APP_NAME', value: appName }
    ]
    scriptContent: '''
set -e
echo "Downloading $PACKAGE_URL..."
curl -L --fail -o /tmp/bot.zip "$PACKAGE_URL"
ls -la /tmp/bot.zip
echo "Waiting for Function App to be ready..."
for i in $(seq 1 30); do
  STATE=$(az functionapp show --resource-group "$RG" --name "$APP_NAME" --query state -o tsv 2>/dev/null || echo Pending)
  echo "  state=$STATE (attempt $i)"
  [ "$STATE" = "Running" ] && break
  sleep 10
done
echo "Deploying package to $APP_NAME..."
az functionapp deployment source config-zip --resource-group "$RG" --name "$APP_NAME" --src /tmp/bot.zip
echo "Done."
'''
  }
}

// --- Outputs ---
output functionAppHostname string = functionApp.properties.defaultHostName
output messagingEndpoint string = 'https://${functionApp.properties.defaultHostName}/api/messages'
output webhookEndpoint string = 'https://${functionApp.properties.defaultHostName}/api/notify'
