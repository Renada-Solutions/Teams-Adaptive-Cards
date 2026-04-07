@description('Base name for all resources (e.g. halopsa-ooh-bot). Must be globally unique as it becomes the Function App hostname.')
param appName string

@description('Azure region for all resources')
param location string = resourceGroup().location

@description('Application (Client) ID from your Azure AD App Registration')
param microsoftAppId string

@secure()
@description('Client secret value from your Azure AD App Registration')
param microsoftAppPassword string

@description('Directory (Tenant) ID from your Azure AD')
param microsoftAppTenantId string

@secure()
@description('Shared secret for authenticating HaloPSA webhook calls. Auto-generated if left blank.')
param notifySecret string = newGuid()

@description('GitHub repository URL containing the bot code')
param repoUrl string = 'https://github.com/YOUR_ORG/YOUR_REPO'

@description('GitHub branch to deploy from')
param repoBranch string = 'main'

// --- Function App Module ---
module functionApp 'modules/functionApp.bicep' = {
  name: 'functionApp'
  params: {
    location: location
    appName: appName
    microsoftAppId: microsoftAppId
    microsoftAppPassword: microsoftAppPassword
    microsoftAppTenantId: microsoftAppTenantId
    notifySecret: notifySecret
    repoUrl: repoUrl
    repoBranch: repoBranch
  }
}

// --- Bot Service Module ---
module botService 'modules/botService.bicep' = {
  name: 'botService'
  params: {
    botName: appName
    microsoftAppId: microsoftAppId
    microsoftAppTenantId: microsoftAppTenantId
    messagingEndpoint: functionApp.outputs.messagingEndpoint
  }
}

// --- Outputs ---
@description('The webhook URL to configure in HaloPSA (POST endpoint)')
output webhookUrl string = functionApp.outputs.webhookEndpoint

@description('The Bot Framework messaging endpoint')
output messagingEndpoint string = functionApp.outputs.messagingEndpoint

@description('The Application (Client) ID — use this in your Teams manifest.json')
output microsoftAppId string = microsoftAppId

@description('Update your Teams manifest.json "id" and "botId" fields with the Application ID above, then zip with icons and upload to Teams.')
output teamsManifestInstructions string = 'Update manifest.json id and botId to "${microsoftAppId}", zip with color.png and outline.png, upload to Teams as a custom app.'
