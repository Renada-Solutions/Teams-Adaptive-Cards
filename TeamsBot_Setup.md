# Teams Bot: OOH Ticket Notification

## Overview

A lightweight Azure-hosted bot that receives webhooks from HaloPSA and sends
Adaptive Cards with @mentions to a Teams channel. Costs effectively $0/month.

```
HaloPSA Runbook
    → POST to Azure Function /api/notify (your webhook endpoint)
    → Azure Function uses Bot Framework to send proactive message
    → Adaptive Card appears in Teams channel with @mentions
```

---

## Architecture

Two Azure resources, both free/near-free:

| Resource | Purpose | Cost |
|---|---|---|
| **Azure Bot** | Bot registration (identity for Teams) | Free |
| **Azure Function App** | Hosts the bot code + webhook endpoint | Consumption plan (~$0/mo) |

The Function App has two HTTP endpoints:

| Endpoint | Purpose |
|---|---|
| `/api/messages` | Bot Framework messaging endpoint (handles bot install events) |
| `/api/notify` | Webhook endpoint that HaloPSA calls to send notifications |

---

## Step-by-Step Setup

### 1. Azure AD App Registration

In Azure Portal → Entra ID → App Registrations:

1. **New registration**
   - Name: `HaloPSA OOH Bot`
   - Supported account types: **Single tenant**
   - Redirect URI: leave blank
2. Note the **Application (client) ID** and **Directory (tenant) ID**
3. **Certificates & secrets** → New client secret → note the **Value**

### 2. Create Azure Bot Resource

In Azure Portal → Create a resource → "Azure Bot":

| Field | Value |
|---|---|
| Bot handle | `halopsa-ooh-bot` |
| Pricing tier | **F0 (Free)** |
| Type of App | Single Tenant |
| App ID | The Application ID from Step 1 |
| App tenant ID | The Tenant ID from Step 1 |
| Messaging endpoint | `https://{your-function-app}.azurewebsites.net/api/messages` |

After creation:
- Go to **Channels** → Add **Microsoft Teams** channel → Save

### 3. Create Azure Function App

In Azure Portal → Create a resource → "Function App":

| Field | Value |
|---|---|
| Runtime stack | **Node.js 20** |
| Hosting plan | **Consumption (Serverless)** |
| Region | Same as your other Azure resources |

After creation, add these **Application Settings** (Configuration → Application settings):

| Name | Value |
|---|---|
| `MicrosoftAppId` | Application ID from Step 1 |
| `MicrosoftAppPassword` | Client secret from Step 1 |
| `MicrosoftAppTenantId` | Tenant ID from Step 1 |
| `ChannelId` | `19:EcDLm480cXMsAuwyrbjWdy_3SWzPMp-fx_RDhkTrdcU1@thread.tacv2` |
| `TeamId` | `6cb5b529-9086-40c2-9d38-ee6b3a04a106` |
| `NOTIFY_SECRET` | A random secret string (HaloPSA must send this to authenticate) |

### 4. Deploy the Bot Code

The Function App needs two functions. See the code files section below.

Deploy via:
- **VS Code** with the Azure Functions extension (easiest)
- **Azure Portal** → Function App → "App files" (for small edits)
- **GitHub Actions** (for CI/CD)

### 5. Install Bot in Teams

1. Create a Teams App package (a .zip containing `manifest.json` + two icon PNGs)
2. In Teams → Apps → **Manage your apps** → **Upload a custom app**
3. Install the app to your team
4. The bot appears in the team — it captures the conversation reference automatically

### 6. Configure HaloPSA

**Custom Integration**:

| Field | Value |
|---|---|
| Name | `OOH Teams Bot` |
| Auth Type | None (auth via secret in body) |
| Base URL | `https://{your-function-app}.azurewebsites.net` |

**Method**:

| Field | Value |
|---|---|
| HTTP Method | POST |
| Endpoint | `/api/notify` |
| Content Type | application/json |

---

## Code Files

### package.json

```json
{
  "name": "halopsa-ooh-bot",
  "version": "1.0.0",
  "dependencies": {
    "botbuilder": "^4.23.0",
    "botframework-connector": "^4.23.0"
  }
}
```

### messages/index.js (Bot Framework endpoint)

This handles bot installation and stores the conversation reference.

```javascript
const {
    BotFrameworkAdapter,
    ConversationReference
} = require('botbuilder');

// Shared state — in production, use Azure Table Storage
// For a single team, app settings work fine
global.conversationReference = global.conversationReference || null;

const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    channelAuthTenant: process.env.MicrosoftAppTenantId
});

adapter.onTurnError = async (context, error) => {
    console.error(`Bot error: ${error}`);
};

module.exports = async function (context, req) {
    await adapter.process(req, context.res, async (turnContext) => {
        // When bot is added to a team, store the conversation reference
        if (turnContext.activity.type === 'conversationUpdate') {
            const ref = TurnContext.getConversationReference(turnContext.activity);
            global.conversationReference = ref;

            // Log it so you can verify
            context.log('Conversation reference stored:', JSON.stringify(ref));
        }

        // Respond to messages (optional — bot doesn't need to chat)
        if (turnContext.activity.type === 'message') {
            await turnContext.sendActivity(
                'I am the OOH notification bot. I post alerts when new tickets are logged out of hours.'
            );
        }
    });
};
```

### messages/function.json

```json
{
    "bindings": [
        {
            "authLevel": "anonymous",
            "type": "httpTrigger",
            "direction": "in",
            "name": "req",
            "methods": ["post"]
        },
        {
            "type": "http",
            "direction": "out",
            "name": "res"
        }
    ]
}
```

### notify/index.js (Webhook endpoint for HaloPSA)

This receives the payload from HaloPSA and sends the proactive message.

```javascript
const {
    BotFrameworkAdapter,
    CardFactory,
    MessageFactory
} = require('botbuilder');

const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    channelAuthTenant: process.env.MicrosoftAppTenantId
});

module.exports = async function (context, req) {
    // Verify the shared secret
    const secret = req.body?.secret;
    if (secret !== process.env.NOTIFY_SECRET) {
        context.res = { status: 401, body: 'Unauthorized' };
        return;
    }

    // Check we have a conversation reference
    if (!global.conversationReference) {
        context.res = {
            status: 400,
            body: 'Bot has not been installed in a team yet. No conversation reference stored.'
        };
        return;
    }

    const { agents, cardJson, ticketId, ticketSummary } = req.body;

    // Parse the adaptive card
    let card;
    try {
        card = typeof cardJson === 'string' ? JSON.parse(cardJson) : cardJson;
    } catch (e) {
        context.res = { status: 400, body: 'Invalid cardJson' };
        return;
    }

    // Build @mention entities
    const mentions = [];
    const mentionTexts = [];

    if (agents && Array.isArray(agents)) {
        agents.forEach((agent, index) => {
            mentions.push({
                type: 'mention',
                text: `<at>${agent.displayName}</at>`,
                mentioned: {
                    id: agent.azureId,
                    name: agent.displayName
                }
            });
            mentionTexts.push(`<at>${agent.displayName}</at>`);
        });
    }

    // Send the proactive message
    try {
        await adapter.continueConversationAsync(
            process.env.MicrosoftAppId,
            global.conversationReference,
            async (turnContext) => {
                const cardAttachment = CardFactory.adaptiveCard(card);

                const activity = MessageFactory.attachment(cardAttachment);
                activity.entities = mentions;

                // The text field with <at> tags triggers the @mention notifications
                if (mentionTexts.length > 0) {
                    activity.text = mentionTexts.join(' ');
                }

                await turnContext.sendActivity(activity);
            }
        );

        context.res = { status: 200, body: { success: true, ticketId } };
    } catch (error) {
        context.log.error('Failed to send proactive message:', error);
        context.res = { status: 500, body: { error: error.message } };
    }
};
```

### notify/function.json

```json
{
    "bindings": [
        {
            "authLevel": "anonymous",
            "type": "httpTrigger",
            "direction": "in",
            "name": "req",
            "methods": ["post"]
        },
        {
            "type": "http",
            "direction": "out",
            "name": "res"
        }
    ]
}
```

### Teams App Manifest (manifest.json)

Create a zip file containing this manifest + two PNG icons (192x192 and 32x32):

```json
{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.17/MicrosoftTeams.schema.json",
    "manifestVersion": "1.17",
    "version": "1.0.0",
    "id": "YOUR_APP_ID_FROM_STEP_1",
    "developer": {
        "name": "Renada",
        "websiteUrl": "https://renada.co.uk",
        "privacyUrl": "https://renada.co.uk/privacy",
        "termsOfUseUrl": "https://renada.co.uk/terms"
    },
    "name": {
        "short": "OOH Alerts",
        "full": "Out-of-Hours Ticket Alerts"
    },
    "description": {
        "short": "Notifies on-call engineers about new OOH tickets",
        "full": "Posts Adaptive Card notifications to the On-Call Alerts channel when new tickets are logged out of hours in HaloPSA."
    },
    "icons": {
        "outline": "outline.png",
        "color": "color.png"
    },
    "accentColor": "#FF0000",
    "bots": [
        {
            "botId": "YOUR_APP_ID_FROM_STEP_1",
            "scopes": ["team"],
            "supportsFiles": false,
            "isNotificationOnly": true
        }
    ],
    "permissions": ["identity", "messageTeamMembers"],
    "validDomains": []
}
```

> Note: `isNotificationOnly: true` — the bot only sends messages, it doesn't need to receive or respond to user messages.

---

## What HaloPSA Sends

```json
{
    "secret": "YOUR_NOTIFY_SECRET",
    "agents": [
        {
            "azureId": "azure-ad-object-id-1",
            "email": "jacob@renada.co.uk",
            "displayName": "Jacob Newman"
        },
        {
            "azureId": "azure-ad-object-id-2",
            "email": "connor@renada.co.uk",
            "displayName": "Connor Fagan"
        }
    ],
    "cardJson": "{THE_FULL_ADAPTIVE_CARD_JSON_AS_A_STRING}",
    "ticketId": "10066",
    "ticketSummary": "Server down - unable to connect to production environment"
}
```

---

## HaloPSA Runbook Flow

```
Step 1: Report    → Get all agents currently on shift
                    (returns: Azure AD object ID, email, display name)

Step 2: Decision  → Any agents returned?
         ├── NO  → End (or fallback)
         └── YES → Continue

Step 3: Build     → Actions array from ticket
                    → <<CF_ACTIONS_JSON!>>

Step 4: Build     → Agents array from Step 1
                    → <<CF_AGENTS_JSON!>>

Step 5: Build     → Full Adaptive Card JSON string
                    → <<CF_CARD_JSON!>>

Step 6: Method    → POST to Azure Function /api/notify
                    Body includes secret, agents, cardJson, ticketId, ticketSummary
```

---

## Testing

1. **Deploy the Function App** with both functions
2. **Create the Teams app package** (manifest.json + icons in a .zip)
3. **Upload to Teams** → Install in your team
4. **Verify bot installation** → Check Function App logs for "Conversation reference stored"
5. **Test with Postman** → POST to `/api/notify` with sample payload
6. **Verify** card appears in channel with @mentions and notifications
7. **Connect HaloPSA** → Configure custom integration → End-to-end test

---

## Conversation Reference Persistence

The code above stores the conversation reference in `global` memory, which resets
when the Azure Function cold-starts. For production reliability, you should persist
it to **Azure Table Storage** (free tier included). However, for initial testing,
the global approach works — just re-add the bot to the team if it stops working
after a cold start.

A simple fix: store the reference in an App Setting or Azure Blob after first
capture, and read it back on cold start.
