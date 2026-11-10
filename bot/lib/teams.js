const { BotFrameworkAdapter, CardFactory, MessageFactory } = require('botbuilder');
const { load } = require('./conversationStore');

const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    channelAuthTenant: process.env.MicrosoftAppTenantId
});

// Verify the NOTIFY_SECRET. Halo runbooks send it as a Bearer token; Halo webhooks may only offer
// Basic auth (secret as username or password) or a URL, so ?token= is accepted too.
function isAuthorised(req) {
    const secret = process.env.NOTIFY_SECRET;
    if (!secret) return false;

    const authHeader = req.headers?.authorization || req.headers?.Authorization || '';
    if (authHeader.startsWith('Bearer ')) return authHeader.slice(7) === secret;
    if (authHeader.startsWith('Basic ')) {
        const decoded = Buffer.from(authHeader.slice(6), 'base64').toString('utf8');
        const sep = decoded.indexOf(':');
        const user = sep === -1 ? decoded : decoded.slice(0, sep);
        const pass = sep === -1 ? '' : decoded.slice(sep + 1);
        return user === secret || pass === secret;
    }
    return req.query?.token === secret;
}

// Build @mentions (users, or Teams tags with type "tag") and inject into the card's msteams.entities.
// The mention TextBlock goes after the element with id "agents-on-shift-label",
// or at the top of the card if that label is not found.
// Drop entries with a blank name or ID (e.g. an unassigned ticket's agent variables),
// since Teams rejects the whole message if a mention has no ID
function validMentions(agents) {
    if (!Array.isArray(agents)) return [];
    return agents.filter(a => a && String(a.displayName || '').trim() && String(a.azureId || '').trim());
}

function addMentions(card, agents) {
    agents = validMentions(agents);
    if (agents.length === 0) return;

    const msteamsEntities = [];
    const mentionTexts = [];

    agents.forEach((agent) => {
        msteamsEntities.push({
            type: 'mention',
            text: `<at>${agent.displayName}</at>`,
            mentioned: {
                id: agent.azureId,
                name: agent.displayName,
                // Teams tag mentions need type "tag" and the tag's full Graph ID (base64) in azureId
                ...(agent.type === 'tag' ? { type: 'tag' } : {})
            }
        });
        mentionTexts.push(`<at>${agent.displayName}</at>`);
    });

    // Add msteams.entities to the card JSON so Teams resolves these as real @mentions
    card.msteams = { ...(card.msteams || {}), entities: msteamsEntities };

    const mentionBlock = {
        type: 'TextBlock',
        text: mentionTexts.join(', '),
        weight: 'Bolder',
        wrap: true,
        spacing: 'None'
    };

    if (!card.body) return;
    for (const element of card.body) {
        if (element.items && Array.isArray(element.items)) {
            const labelIndex = element.items.findIndex(i => i.id === 'agents-on-shift-label');
            if (labelIndex !== -1) {
                element.items.splice(labelIndex + 1, 0, mentionBlock);
                return;
            }
        }
    }
    // Fallback: prepend to card body if label not found
    card.body.unshift(mentionBlock);
}

// Teams channel IDs look like 19:xxxx@thread.tacv2 (from the channel's "Get link" URL, URL-decoded)
const CHANNEL_ID_PATTERN = /^19:[^;\s]+@thread\.(tacv2|skype)$/;

function isValidChannelId(channelId) {
    return CHANNEL_ID_PATTERN.test(channelId);
}

// Run logic against the stored conversation, or another conversation (channel or thread) in the same team.
// Returns false if the bot has not been installed yet.
async function withConversation(conversationId, logic) {
    const storedReference = await load();
    if (!storedReference) return false;

    // Copy rather than mutate: load() returns a cached object shared across invocations
    const conversationReference = conversationId
        ? { ...storedReference, conversation: { ...storedReference.conversation, id: conversationId } }
        : storedReference;

    await adapter.continueConversation(conversationReference, logic);
    return true;
}

// Send a card to the stored conversation, or to another channel in the same team if channelId is given.
// Returns { conversationId, activityId } so the message can be updated later, or null if the bot is not installed.
async function sendCard(card, channelId) {
    let result = null;
    const installed = await withConversation(channelId, async (turnContext) => {
        const activity = MessageFactory.attachment(CardFactory.adaptiveCard(card));
        const response = await turnContext.sendActivity(activity);
        result = {
            conversationId: turnContext.activity.conversation.id,
            activityId: response && response.id
        };
    });
    return installed ? result : null;
}

// Replace a card the bot sent earlier
async function updateCard(conversationId, activityId, card) {
    return withConversation(conversationId, async (turnContext) => {
        const activity = MessageFactory.attachment(CardFactory.adaptiveCard(card));
        activity.id = activityId;
        await turnContext.updateActivity(activity);
    });
}

// Reply in the thread under a message the bot sent earlier
async function replyInThread(conversationId, activityId, text) {
    return withConversation(`${conversationId};messageid=${activityId}`, async (turnContext) => {
        await turnContext.sendActivity(MessageFactory.text(text));
    });
}

module.exports = { isAuthorised, isValidChannelId, validMentions, addMentions, sendCard, updateCard, replyInThread };
