# HaloPSA Teams Notification Bot

A Microsoft Teams bot that sends Adaptive Card notifications when new tickets are logged in HaloPSA outside of business hours. On-call engineers receive @mentions with full ticket details, SLA targets, and quick-action buttons.

## Features

- Adaptive Card notifications with full ticket context
- @mentions for on-shift engineers (Teams push notifications)
- Conditional VIP customer badges
- HTML-to-markdown conversion for ticket descriptions
- SLA response/resolution target display
- Quick-action buttons (Triage, Verify User, Cancel, etc.)
- Bearer token authentication for webhook security

## Architecture

```
HaloPSA Runbook → POST /api/notify → Azure Function App → Teams Channel
```

**Azure resources created:**
- Storage Account (required by Functions)
- App Service Plan (Consumption/Y1 — free tier)
- Function App (Node.js 20)
- Azure Bot Service (F0 — free tier)
- Teams Channel on the Bot

---

## Quick Deploy

### Step 1: Create an Azure AD App Registration

This must be done before deploying — ARM templates cannot create app registrations with client secrets.

1. Go to **Azure Portal** → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Set the name to something like `HaloPSA OOH Bot`
3. Set **Supported account types** to **Accounts in this organizational directory only (Single tenant)**
4. Click **Register**
5. On the overview page, copy:
   - **Application (client) ID** — you'll need this as `microsoftAppId`
   - **Directory (tenant) ID** — you'll need this as `microsoftAppTenantId`
6. Go to **Certificates & secrets** → **Client secrets** → **New client secret**
7. Set a description (e.g. `bot-secret`) and expiry, then click **Add**
8. **Copy the secret Value immediately** (it won't be shown again) — you'll need this as `microsoftAppPassword`

### Step 2: Deploy to Azure

Click the button below to deploy all Azure resources automatically:

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2FYOUR_ORG%2FYOUR_REPO%2Fmain%2Fazuredeploy.json)

> **Note:** Update the button URL above with your actual GitHub repository URL after pushing to GitHub.

Fill in the deployment form:

| Parameter | Description |
|-----------|-------------|
| **App Name** | Globally unique name (e.g. `mycompany-ooh-bot`). Becomes the Function App hostname. |
| **Location** | Azure region (e.g. `UK South`, `East US`) |
| **Microsoft App Id** | Application (Client) ID from Step 1 |
| **Microsoft App Password** | Client secret value from Step 1 |
| **Microsoft App Tenant Id** | Directory (Tenant) ID from Step 1 |
| **Notify Secret** | Leave blank to auto-generate, or set your own Bearer token |
| **Repo Url** | Your GitHub repo URL (pre-filled) |
| **Repo Branch** | Branch to deploy from (default: `main`) |

Click **Review + create** → **Create**. Deployment takes 3-5 minutes.

### Step 3: Get Deployment Outputs

After deployment completes, go to **Resource Group** → **Deployments** → click the deployment → **Outputs**:

| Output | Use it for |
|--------|-----------|
| **webhookUrl** | HaloPSA Custom Integration endpoint URL |
| **microsoftAppId** | Teams manifest `id` and `botId` fields |
| **teamsManifestInstructions** | Summary of Teams app setup steps |

To retrieve the `NOTIFY_SECRET` (if auto-generated): go to the **Function App** → **Configuration** → **Application settings** → **NOTIFY_SECRET**.

### Step 4: Create & Install the Teams App

1. Edit [bot/manifest/manifest.json](bot/manifest/manifest.json):
   - Set `"id"` to your **Application (Client) ID**
   - Set `"bots"[0]."botId"` to the same ID
   - Update `"validDomains"` to your Function App hostname (e.g. `mycompany-ooh-bot.azurewebsites.net`)
   - Update `"developer"` fields with your company info
2. Zip the following files together:
   - `manifest.json`
   - `color.png`
   - `outline.png`
3. In Teams: **Apps** → **Manage your apps** → **Upload an app** → **Upload a custom app**
4. Install the app to the team/channel where you want notifications
5. The bot automatically stores the conversation reference when installed

### Step 5: Configure HaloPSA

Set up a Custom Integration in HaloPSA:

- **URL:** The `webhookUrl` from deployment outputs (e.g. `https://mycompany-ooh-bot.azurewebsites.net/api/notify`)
- **Auth Type:** Bearer Token
- **Token:** The `NOTIFY_SECRET` value from Function App configuration
- **Method:** POST
- **Body:** See [test-payload.json](test-payload.json) for the expected format

The POST body should include:
```json
{
    "ticketId": "<<ticket^id!>>",
    "ticketSummary": "<<ticket^summary!>>",
    "agents": [
        {
            "displayName": "Agent Name",
            "azureId": "their-azure-ad-object-id"
        }
    ],
    "cardJson": { }
}
```

The `cardJson` field should contain a fully populated Adaptive Card — see [OOHTicketNotification.json](OOHTicketNotification.json) for the template with HaloPSA variables.

---

## Manual Deployment

For step-by-step manual deployment instructions (without the ARM template), see [TeamsBot_Setup.md](TeamsBot_Setup.md).

---

## Testing

Send a test notification using curl:

```bash
curl -X POST https://YOUR-APP.azurewebsites.net/api/notify \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer YOUR_NOTIFY_SECRET" \
  -d @test-payload.json
```

## Project Structure

```
├── azuredeploy.json              # ARM template (Deploy to Azure button)
├── .deployment                   # Tells Kudu to deploy from bot/ subfolder
├── infra/                        # Bicep source templates
│   ├── main.bicep
│   └── modules/
│       ├── functionApp.bicep
│       └── botService.bicep
├── bot/                          # Azure Function App code
│   ├── host.json
│   ├── package.json
│   ├── messages/                 # Bot installation endpoint
│   │   ├── function.json
│   │   └── index.js
│   ├── notify/                   # HaloPSA webhook endpoint
│   │   ├── function.json
│   │   └── index.js
│   └── manifest/                 # Teams app package
│       ├── manifest.json
│       ├── color.png
│       └── outline.png
├── OOHTicketNotification.json    # Adaptive Card template (HaloPSA variables)
└── test-payload.json             # Sample test payload
```

## License

Internal use — Renada Solutions Ltd.
