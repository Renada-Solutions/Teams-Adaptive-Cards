# Power Automate Flow: OOH Ticket Notification

## Overview

HaloPSA sends a webhook to Power Automate with ticket data and on-call agent info.
Power Automate posts the Adaptive Card to a Teams channel with @mentions for all agents.

---

## Prerequisites

- A **service account** (e.g., `svc-halo-notify@renada.co.uk`) that:
  - Has a Power Automate license (or use an existing licensed account)
  - Is a member of the "On-Call Alerts" Teams channel
  - Owns the Teams connector connection in the flow
- The flow runs under this account's Teams connection (delegated permissions)

---

## Flow Steps

### Step 1: Trigger — "When an HTTP request is received"

**Type**: Instant cloud flow trigger

**HTTP Method**: POST

**Request Body JSON Schema**:
```json
{
    "type": "object",
    "properties": {
        "agents": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "email": { "type": "string" },
                    "displayName": { "type": "string" }
                }
            }
        },
        "cardJson": { "type": "string" },
        "ticketId": { "type": "string" },
        "ticketSummary": { "type": "string" }
    }
}
```

> After saving, Power Automate generates the **HTTP POST URL** — this is what you configure in HaloPSA as the webhook endpoint.

---

### Step 2: Initialize Variable — "MentionsString"

**Type**: Initialize variable

| Field | Value |
|---|---|
| Name | `MentionsString` |
| Type | String |
| Value | *(empty)* |

This will accumulate all the `<at>` mention tokens.

---

### Step 3: For Each — Loop through agents

**Type**: Apply to each

**Input**: `@{triggerBody()?['agents']}`

#### Step 3a: Get @mention token for a user

**Type**: Microsoft Teams → "Get an @mention token for a user"

| Field | Value |
|---|---|
| User | `@{items('For_each')?['email']}` |

**Output**: The mention token (e.g., `<at>Jacob Newman</at>`)

#### Step 3b: Append to string variable

**Type**: Append to string variable

| Field | Value |
|---|---|
| Name | `MentionsString` |
| Value | `@{body('Get_@mention_token_for_a_user')?['atMention']} ` |

> Note the trailing space — separates multiple mentions.

---

### Step 4: Post adaptive card in a chat or channel

**Type**: Microsoft Teams → "Post adaptive card in a chat or channel"

| Field | Value |
|---|---|
| Post as | Flow bot |
| Post in | Channel |
| Team | *(select your team)* |
| Channel | On-Call Alerts |
| Adaptive Card | `@{triggerBody()?['cardJson']}` |
| Message | `@{variables('MentionsString')}` |

> The **Message** field is where the @mention tokens go. This is the text that appears above the card and triggers the Teams notifications for each mentioned agent.

---

## What HaloPSA Sends

Your HaloPSA runbook's final step POSTs this to the Power Automate webhook URL:

```json
{
    "agents": [
        {
            "email": "jacob@renada.co.uk",
            "displayName": "Jacob Newman"
        },
        {
            "email": "connor@renada.co.uk",
            "displayName": "Connor Fagan"
        }
    ],
    "cardJson": "{THE_FULL_ADAPTIVE_CARD_JSON_AS_A_STRING}",
    "ticketId": "10066",
    "ticketSummary": "Server down - unable to connect to production environment"
}
```

> **cardJson**: This is `OOHTicketNotification.json` with all `<<variable!>>` tokens already resolved by HaloPSA, then JSON-stringified.

---

## HaloPSA Custom Integration

| Field | Value |
|---|---|
| Name | `Power Automate OOH Webhook` |
| Auth Type | None |
| Base URL | *(leave empty or use the full URL in the method)* |

### Method: "Send OOH Notification"

| Field | Value |
|---|---|
| HTTP Method | POST |
| Endpoint | The full Power Automate HTTP trigger URL |
| Content Type | application/json |
| Request Body | The JSON payload above (with HaloPSA variables) |

---

## Updated HaloPSA Runbook Flow

```
Step 1: Report    → Get all agents currently on shift
                    (returns: email, display name, Halo ID, etc.)

Step 2: Decision  → Any agents returned?
         ├── NO  → End (or fallback notification)
         └── YES → Continue

Step 3: Build     → Construct actions array from ticket's available actions
                    → <<CF_ACTIONS_JSON!>>

Step 4: Build     → Construct agents array from Step 1 results
                    → <<CF_AGENTS_JSON!>>
                    e.g. [{"email":"jacob@renada.co.uk","displayName":"Jacob Newman"}, ...]

Step 5: Build     → Construct full Adaptive Card JSON string
                    → <<CF_CARD_JSON!>>
                    (OOHTicketNotification.json with all variables resolved)

Step 6: Method    → POST to Power Automate webhook URL
                    Body: { "agents": <<CF_AGENTS_JSON!>>,
                            "cardJson": <<CF_CARD_JSON!>>,
                            "ticketId": "<<ticket^id!>>",
                            "ticketSummary": "<<ticket^summary!>>" }
```

---

## Testing

1. **Save the Power Automate flow** → copy the HTTP POST URL
2. **Test with Postman/curl** → POST the sample payload to the URL
3. **Verify** the card appears in the On-Call Alerts channel with @mentions
4. **Verify** each mentioned agent receives a Teams notification
5. **Connect HaloPSA** → configure the custom integration with the webhook URL
6. **End-to-end test** → log a test OOH ticket → confirm full flow

---

## Notes

- The Power Automate webhook URL contains a SAS token — treat it like a secret
- The flow runs under the service account's Teams connection — if that account is disabled, the flow breaks
- Power Automate has a 100-action run limit on free plans — the For Each loop counts per agent, so this is fine for typical on-call rosters
- If the webhook URL changes (flow re-created), update the HaloPSA integration method
