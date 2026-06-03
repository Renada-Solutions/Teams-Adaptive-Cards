const { BotFrameworkAdapter, TurnContext } = require('botbuilder');
const { save } = require('../lib/conversationStore');

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
                    const ref = TurnContext.getConversationReference(turnContext.activity);
                    await save(ref);
                    context.log('Bot installed — conversation reference saved to blob storage');
                    context.log(JSON.stringify(ref, null, 2));
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
