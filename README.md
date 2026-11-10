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
- Function App (Linux Flex Consumption, Node.js 22)
- Azure Bot Service (F0 — free tier)
- Teams Channel on the Bot

---

## Quick Deploy

### Step 1: Create an Azure AD App Registration

This must be done before deploying — ARM templates cannot create app registrations with client secrets.

1. Go to **Azure Portal** → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Set the name to something like `HaloXTeams`
3. Set **Supported account types** to **Accounts in this organizational directory only (Single tenant)**
4. Click **Register**
5. On the overview page, copy:
   - **Application (client) ID** — you'll need this as `microsoftAppId`
   - **Directory (tenant) ID** — you'll need this as `microsoftAppTenantId`
6. Go to **Certificates & secrets** → **Client secrets** → **New client secret**
7. Set a description (e.g. `bot-secret`) and expiry, then click **Add**
8. **Copy the secret Value immediately** (it won't be shown again) — you'll need this as `microsoftAppPassword`

### Step 2: Deploy Infrastructure to Azure

Click the button below. ARM provisions all the resources (Flex Consumption Function App, Bot Service, Teams channel, storage). The Function App is empty after this — code is pushed in Step 3.

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2FRenada-Solutions%2FTeams-Adaptive-Cards%2Fmain%2Fazuredeploy.json)

Fill in the deployment form:

| Parameter | Description |
|-----------|-------------|
| **App Name** | Globally unique name (e.g. `haloxteams`). Becomes the Function App hostname. Use only lowercase letters, digits, and hyphens. |
| **Microsoft App Id** | Application (Client) ID from Step 1 |
| **Microsoft App Password** | Client secret value from Step 1 |

Click **Review + create** → **Create**. Deployment takes ~1-2 minutes.

After deployment, go to **Resource Group** → **Deployments** → click the deployment → **Outputs**. Copy `deployCommand` — you'll paste it in Step 3.

| Output | Use it for |
|--------|-----------|
| **deployCommand** | The command you paste into Cloud Shell in Step 3 |
| **webhookUrl** | HaloPSA Custom Integration endpoint URL |
| **microsoftAppId** | Teams manifest `id` and `botId` fields |

To retrieve the auto-generated `NOTIFY_SECRET`: **Function App** → **Configuration** → **Application settings** → **NOTIFY_SECRET**.

### Step 3: Deploy the Bot Code

Flex Consumption requires code to be uploaded to its deployment blob container — ARM templates can't do this themselves. One copy-paste does it:

1. Open [Azure Cloud Shell](https://shell.azure.com) (Bash). It's browser-based and already authenticated as you — no install needed.
2. Paste the `deployCommand` value from the deployment outputs and hit Enter. It looks like:

   ```bash
   curl -L -o bot.zip https://github.com/Renada-Solutions/Teams-Adaptive-Cards/releases/latest/download/bot.zip && az functionapp deployment source config-zip --resource-group <RG> --name <APP> --src bot.zip
   ```

3. Wait ~60 seconds. When `az` finishes, refresh the Function App overview — `messages` and `notify` will appear and the runtime version will go from `Error` to a real version.

Re-run the same command anytime to pull the latest bot release. No infrastructure redeploy required.

### Step 4: Create & Install the Teams App

1. Edit [bot/manifest/manifest.json](bot/manifest/manifest.json):
   - Set `"id"` to your **Application (Client) ID**
   - Set `"bots"[0]."botId"` to the same ID
   - Update `"validDomains"` to your Function App hostname (e.g. `haloxteams.azurewebsites.net`)
   - Update `"developer"` fields with your company info
2. Zip the following files together:
   - `manifest.json`
   - `color.png`
   - `outline.png`
3. In Teams: **Apps** → **Manage your apps** → **Upload an app** → **Upload a custom app**
4. Install the app to the team/channel where you want notifications
5. The bot automatically stores the conversation reference when installed

### Step 4: Configure HaloPSA

Set up a Custom Integration in HaloPSA:

- **URL:** The `webhookUrl` from deployment outputs (e.g. `https://haloxteams.azurewebsites.net/api/notify`)
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

### Live Chat notifications

Add an **Execute a Runbook** step to the chat profile flow (after the initial questions, before Start Live Chat). The runbook POSTs to `/api/chat` with the same Bearer token. The bot builds the card itself, so no `cardJson` is needed.

The body can be either:

- The raw Halo chat object as-is, or
- A wrapper, if the runbook looks up extra details first:

```json
{
    "chat": { },
    "ticket": { "id": 2146689, "summary": "", "client": "", "status": "", "priority": "", "agent": "" },
    "agents": [ { "displayName": "Agent Name", "azureId": "their-azure-ad-object-id" } ]
}
```

**Channel:** chat cards go to the `CHAT_CHANNEL_ID` app setting (or the `chatChannelId` deployment parameter) if set, otherwise to the channel the bot was installed to. A `channelId` in the body overrides both. To get the ID: in Teams, open the channel's **...** menu > **Get link to channel**, then URL-decode the `19%3a...%40thread.tacv2` part to `19:...@thread.tacv2`. The channel must be in a team the bot is installed in.

**Tag mentions:** to @mention a Teams tag (e.g. a support team), add it to `agents` with `"type": "tag"` and the tag's full Graph ID in `azureId`: `{ "displayName": "Level 1 Support Team", "azureId": "<tag id>", "type": "tag" }`. Use the `id` exactly as Graph returns it from `GET https://graph.microsoft.com/v1.0/teams/{team-id}/tags` (or the Power Automate Teams action "List all tags for a team", which needs no admin consent). It is base64 of `{tenantId}##{teamId}##{shortTagId}`. Tag mentions only work in standard channels, not private or shared ones.

All `ticket` fields and `agents` are optional; fields that arrive as the text `null` are ignored. The ticket shown on the card is picked in this order: a numeric `ticket.id` from the runbook, then `ticket_id` on the chat, then the last end-user message that is only a number (flagged on the card as not validated). It is labelled "existing ticket" if the user typed its number in the chat flow, otherwise "new ticket raised"; send `ticket.isNew` (true/false) to override. The `TICKET_PREFIX` app setting sets how ticket numbers appear (e.g. `VT:` gives `VT:2148969`; default `#`). Links use the origin of the chat's `user_chat_url`, or the `HALO_URL` app setting if set.

**Agent joined:** when the card is sent, the bot saves where it went (`botstate/chats/<chatId>.json`). Add a Halo **Webhook** (Configuration > Integrations > Webhooks) on the **Chat Message Sent** event, message types **Messages from agents**, filtered to this chat profile, pointing at the same `/api/chat` endpoint, with Basic authentication (password = `NOTIFY_SECRET`). Halo sends a small envelope (`event: "chat message sent"`, `object_id` = chat ID, `agent_id`) with no agent name, so the card shows "Agent joined" without a name. A runbook can instead send:

```json
{ "event": "agentJoined", "chatId": 1982, "agentName": "Agent Name" }
```

If the webhook body also carries the full chat as `chatDetails` (either alongside the envelope fields, or as `{ "request": <envelope>, "chatDetails": <chat> }`) (as the HaloPSA webhook can be set up to send), the agent's name comes from its `participants`, and "Chat so far" is rebuilt from it every time a newer copy arrives (by `last_message_id`): flow answers, live messages from the user and agents (agents highlighted), and system events such as "joined the chat". Only the last 20 entries are shown.

"Chat so far" is posted as a single reply in the card's thread when the card is sent, then edited in place rather than adding a reply per message, so thread followers (including anyone @mentioned on the card) aren't notified for every chat message. The card's start time uses Adaptive Card `DATE()`/`TIME()` functions, so each viewer sees it in their own time zone.

The bot edits the original card to show "Agent joined: Agent Name" with a green header. If the edit fails it replies in the card's thread instead. Only the first agent message per chat does anything; user messages (`agent_id` 0) and chats with no card are ignored with a 200 so Halo doesn't count them as failures.

**Management escalation:** add `escalation` to the start runbook's body to escalate when no agent responds in time. The threshold can come from a Halo variable, so it can differ by client or priority:

```json
"escalation": {
    "afterMinutes": 5,
    "agents": [ { "displayName": "UKTSMgmt", "azureId": "<tag id>", "type": "tag" } ]
}
```

The `escalate` timer function runs every minute. If no agent message has arrived `afterMinutes` after the card was posted, it posts a reply in the card's thread @mentioning `agents` and turns the card's header red ("No response after 5 min - escalated"). An agent responding later turns it green as usual. Each chat is escalated at most once; an agent responding (or the chat closing, if the webhook sends all messages) cancels it. Without `escalation` nothing changes.

**Installing in more teams:** the bot keeps one stored conversation reference, replaced each time it is installed in a team. Set `NOTIFY_CHANNEL_ID` (OOH alerts) and `CHAT_CHANNEL_ID` (chat alerts) so a new install doesn't move either. Tag mentions only resolve when the bot is installed in the team that owns the tags.

**Auth:** every endpoint accepts `NOTIFY_SECRET` as a Bearer token, as Basic auth (username or password), or as `?token=` on the URL, since Halo webhooks may not offer a Bearer option.

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
├── azuredeploy.json              # ARM template — infra + code-pull (Deploy to Azure button)
├── .github/workflows/
│   └── build-release-bot.yml     # GitHub Actions: builds bot.zip and publishes it as a "latest" release
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
│   ├── chat/                     # HaloPSA Live Chat runbook/webhook endpoint
│   │   ├── function.json
│   │   └── index.js
│   ├── escalate/                 # Timer (every minute): management escalation for unanswered chats
│   │   ├── function.json
│   │   └── index.js
│   ├── lib/                      # Shared: conversation/chat stores, chat cards, auth, Teams send
│   └── manifest/                 # Teams app package
│       ├── manifest.json
│       ├── color.png
│       └── outline.png
├── OOHTicketNotification.json    # Adaptive Card template (HaloPSA variables)
└── test-payload.json             # Sample test payload
```

## Card Templates

| File | Purpose |
|------|---------|
| `OOHTicketNotification.json` | **Production template.** Adaptive Card with HaloPSA `<<variable!>>` syntax — paste directly into HaloPSA Custom Integration as the card body. |
| `OOHTicketNotification_Designer.json` | Template for the [Teams Adaptive Card Designer](https://adaptivecards.io/designer/). Uses `${variable}` syntax for live preview in the designer tool. |
| `OOHTicketNotification_SampleData.json` | Sample data file for use alongside `OOHTicketNotification_Designer.json` in the Adaptive Card Designer. |
| `TicketNotificationFilled.json` | Fully populated example card — useful for understanding the rendered output without needing a live HaloPSA instance. |
| `OOHTicketNotification_TeamsWebhook.json` | Variant formatted for direct Teams Incoming Webhook (no bot required). Use this if you want to send cards without the Azure Bot infrastructure. |
| `MSTeamsCardTemplate.json` | Legacy generic reference card. Kept for historical reference. |
| `MSTeamsCardTemplateCIPP.json` | CIPP-specific variant card template. Kept for reference if adapting this bot for CIPP integrations. |

---

## License

Internal use — Renada Solutions Ltd.
