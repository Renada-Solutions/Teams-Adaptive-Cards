const { BotFrameworkAdapter, TurnContext } = require('botbuilder');
const fs = require('fs');
const path = require('path');

// Persist conversation reference to file so it survives cold starts
const refPath = process.env.HOME
    ? path.join(process.env.HOME, 'data', 'conversationReference.json')
    : path.join(__dirname, '..', 'conversationReference.json');

function loadConversationReference() {
    try {
        if (fs.existsSync(refPath)) {
            return JSON.parse(fs.readFileSync(refPath, 'utf8'));
        }
    } catch (e) {
        console.error('Failed to load conversation reference:', e.message);
    }
    return null;
}

function saveConversationReference(ref) {
    try {
        const dir = path.dirname(refPath);
        if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true });
        }
        fs.writeFileSync(refPath, JSON.stringify(ref, null, 2));
    } catch (e) {
        console.error('Failed to save conversation reference:', e.message);
    }
}

// Load on startup
if (!global.conversationReference) {
    global.conversationReference = loadConversationReference();
}

const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    channelAuthTenant: process.env.MicrosoftAppTenantId
});

adapter.onTurnError = async (context, error) => {
    console.error(`Bot error: ${error.message}`);
    await context.sendActivity('An error occurred. Check the Function App logs.');
};

module.exports = async function (context, req) {
    context.log('Messages endpoint hit');

    await adapter.process(req, context.res, async (turnContext) => {
        // Bot added to team — store the conversation reference
        if (turnContext.activity.type === 'conversationUpdate') {
            const membersAdded = turnContext.activity.membersAdded || [];
            const botId = turnContext.activity.recipient.id;

            for (const member of membersAdded) {
                if (member.id === botId) {
                    global.conversationReference = TurnContext.getConversationReference(turnContext.activity);
                    saveConversationReference(global.conversationReference);
                    context.log('Bot installed — conversation reference stored and persisted');
                    context.log(JSON.stringify(global.conversationReference, null, 2));
                }
            }
        }

        // Someone messages the bot directly
        if (turnContext.activity.type === 'message') {
            await turnContext.sendActivity(
                'I am the OOH notification bot. I post alerts when new tickets are logged out of hours.'
            );
        }
    });
};
