const { BotFrameworkAdapter, CardFactory, MessageFactory } = require('botbuilder');
const { load } = require('./conversationStore');

const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    channelAuthTenant: process.env.MicrosoftAppTenantId
});

// Verify authorization via Bearer token (configured on HaloPSA Custom Integration / runbook)
function isAuthorised(req) {
    const authHeader = req.headers?.authorization || req.headers?.Authorization || '';
    const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : '';
    return token !== '' && token === process.env.NOTIFY_SECRET;
}

// Build @mentions and inject into the card's msteams.entities.
// The mention TextBlock goes after the element with id "agents-on-shift-label",
// or at the top of the card if that label is not found.
function addMentions(card, agents) {
    if (!agents || !Array.isArray(agents) || agents.length === 0) return;

    const msteamsEntities = [];
    const mentionTexts = [];

    agents.forEach((agent) => {
        msteamsEntities.push({
            type: 'mention',
            text: `<at>${agent.displayName}</at>`,
            mentioned: {
                id: agent.azureId,
                name: agent.displayName
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

// Send a card to the stored conversation, or to another channel in the same team if channelId is given.
// Returns false if the bot has not been installed yet.
async function sendCard(card, channelId) {
    const storedReference = await load();
    if (!storedReference) return false;

    // Copy rather than mutate: load() returns a cached object shared across invocations
    const conversationReference = channelId
        ? { ...storedReference, conversation: { ...storedReference.conversation, id: channelId } }
        : storedReference;

    await adapter.continueConversation(
        conversationReference,
        async (turnContext) => {
            const cardAttachment = CardFactory.adaptiveCard(card);
            const activity = MessageFactory.attachment(cardAttachment);
            await turnContext.sendActivity(activity);
        }
    );
    return true;
}

module.exports = { isAuthorised, isValidChannelId, addMentions, sendCard };
