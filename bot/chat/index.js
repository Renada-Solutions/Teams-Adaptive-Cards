const {
    isAuthorised, isValidChannelId, isTagMentionError,
    sendCard, updateCard, replyInThread, sendCardInThread, updateCardInThread
} = require('../lib/teams');
const { getChat, saveChat, addPending, removePending } = require('../lib/chatStore');
const {
    agentNameFrom, buildTranscriptCard, cardFromPayload, buildEscalationResolvedCard, resolveTicket
} = require('../lib/chatCard');

// Optional escalation from the start runbook: { afterMinutes: 1-1440, agents: [mentions] }.
// Returns null if absent; { warning } if present but unusable (e.g. a blank Halo variable),
// which skips escalation but still posts the chat card.
function parseEscalation(escalation) {
    if (!escalation) return null;
    const afterMinutes = Number(escalation.afterMinutes);
    if (!(afterMinutes >= 1 && afterMinutes <= 1440)) {
        return { warning: `Escalation skipped: afterMinutes "${escalation.afterMinutes}" is not a number from 1 to 1440` };
    }
    return { afterMinutes, agents: Array.isArray(escalation.agents) ? escalation.agents : [] };
}

// Halo "Chat Message Sent" webhook: { event: "chat message sent", object_id: <chat id>, agent_id, ... }
function isHaloChatWebhook(body) {
    return typeof body.event === 'string' && body.event.toLowerCase() === 'chat message sent' && body.object_id !== undefined;
}

// A chat message after the card was sent. Marks the agent as joined on the first agent message,
// and refreshes "Chat so far" whenever the call carries a newer copy of the chat (chatDetails).
// Calls that don't apply return 200 with "ignored" so Halo doesn't treat them as webhook failures.
async function handleChatMessage(context, body) {
    // Halo's webhook envelope carries the chat ID as object_id
    const chatId = body.chatId || body.chat_id || (isHaloChatWebhook(body) ? body.object_id : undefined);
    if (!chatId) {
        context.res = { status: 400, body: { error: 'agentJoined needs chatId' } };
        return;
    }

    const latestChat = body.chatDetails && Array.isArray(body.chatDetails.messages) ? body.chatDetails : null;
    const isAgentMessage = body.event === 'agentJoined' || Number(body.agent_id) > 0;
    const agentName = String(
        body.agentName || body.agent_name || agentNameFrom(latestChat, body.agent_id) || body.display_name || ''
    ).trim();

    if (!isAgentMessage && !latestChat) {
        context.res = { status: 200, body: { success: true, chatId, ignored: 'not an agent message' } };
        return;
    }

    // A closed or abandoned chat no longer needs escalating (only seen if the webhook sends all messages)
    const isClosed = !!latestChat && !!(latestChat.is_closed || latestChat.is_abandoned);

    // Read-modify-write with an ETag; if another message saved first, re-read and try again
    let record = null;
    let joinsNow = false;
    let isNewer = false;
    let closesNow = false;
    for (let attempt = 0; attempt < 3 && !record; attempt++) {
        const found = await getChat(chatId);
        if (!found) {
            context.log(`Message for chat ${chatId} ignored: no card was sent for it`);
            context.res = { status: 200, body: { success: true, chatId, ignored: 'no card sent for this chat' } };
            return;
        }

        const candidate = found.record;
        joinsNow = isAgentMessage && !candidate.joinedAt;
        isNewer = !!latestChat && Number(latestChat.last_message_id || 0) > Number(candidate.lastMessageId || 0);
        closesNow = isClosed && !candidate.closedAt;
        if (!joinsNow && !isNewer && !closesNow) {
            context.res = { status: 200, body: { success: true, chatId, alreadyUpdated: true } };
            return;
        }

        if (closesNow) candidate.closedAt = new Date().toISOString();
        if (joinsNow) {
            candidate.joinedAt = new Date().toISOString();
            candidate.joinedAgent = agentName;
        }
        if (isNewer) {
            candidate.payload.chat = latestChat;
            candidate.lastMessageId = Number(latestChat.last_message_id || 0);
        }
        if (await saveChat(chatId, candidate, found.etag)) record = candidate;
    }

    if (!record) {
        context.res = { status: 200, body: { success: true, chatId, ignored: 'busy: another update for this chat is in progress' } };
        return;
    }

    const result = { success: true, chatId };

    // An agent responding (or the chat ending) cancels any pending management escalation
    if (joinsNow || closesNow) {
        try {
            await removePending(chatId);
        } catch (error) {
            context.log.warn(`Chat ${chatId}: pending escalation not removed (the timer will skip it anyway): ${error.message}`);
        }
    }

    // First agent message: the main card's header changes to "Agent joined"
    if (joinsNow) {
        result.agentJoined = true;
        try {
            await updateCard(record.conversationId, record.activityId, cardFromPayload(record.payload, record));
            result.card = 'updated';
        } catch (error) {
            context.log.warn(`Card update failed for chat ${chatId}, replying in thread instead: ${error.message}`);
            await replyInThread(record.conversationId, record.activityId, `${agentName || 'An agent'} has joined the chat.`);
            result.card = 'replied';
        }

        // Already escalated: turn the red escalation reply green so the thread doesn't end on it
        if (record.escalatedAt && record.escalationActivityId) {
            try {
                await updateCardInThread(record.conversationId, record.activityId, record.escalationActivityId, buildEscalationResolvedCard(record));
                result.escalation = 'resolved';
            } catch (error) {
                context.log.warn(`Chat ${chatId}: escalation reply not updated: ${error.message}`);
            }
        }
    }

    // Newer chat: refresh the "Chat so far" reply in the thread. A failure here is not worth
    // failing the webhook over; the next message will try again.
    if (isNewer) {
        try {
            result.transcript = await upsertTranscript(chatId, record);
        } catch (error) {
            context.log.warn(`Transcript update failed for chat ${chatId}: ${error.message}`);
            result.warning = `Transcript not updated: ${error.message}`;
        }
    }

    context.log(`Chat ${chatId}: card ${result.card || 'unchanged'}, transcript ${result.transcript || 'unchanged'}${joinsNow ? ` (agent joined: ${agentName || 'name not supplied'})` : ''}`);
    context.res = { status: 200, body: result };
}

// Edit the transcript reply, or post it if the chat doesn't have one yet
async function upsertTranscript(chatId, record) {
    const card = buildTranscriptCard(record.payload.chat);
    if (!card) return 'empty';
    if (record.transcriptActivityId) {
        await updateCardInThread(record.conversationId, record.activityId, record.transcriptActivityId, card);
        return 'updated';
    }
    const replyId = await sendCardInThread(record.conversationId, record.activityId, card);
    await rememberTranscriptReply(chatId, replyId);
    return 'posted';
}

// Save the transcript reply's ID onto the chat record (re-reading if another update saved first)
async function rememberTranscriptReply(chatId, replyId) {
    if (!replyId) return;
    for (let attempt = 0; attempt < 3; attempt++) {
        const found = await getChat(chatId);
        if (!found || found.record.transcriptActivityId) return;
        found.record.transcriptActivityId = replyId;
        if (await saveChat(chatId, found.record, found.etag)) return;
    }
}

module.exports = async function (context, req) {
    context.log('Chat endpoint hit');

    if (!isAuthorised(req)) {
        context.res = { status: 401, body: { error: 'Unauthorized' } };
        return;
    }

    // Halo webhooks may send a JSON string or a one-item array; normalise to an object
    let body = req.body || {};
    if (typeof body === 'string') {
        try { body = JSON.parse(body); } catch { body = {}; }
    }
    if (Array.isArray(body)) body = body[0] || {};

    // Halo webhook with a custom body: { request: <Halo's webhook envelope>, chatDetails: <the chat> }
    if (body.request && typeof body.request === 'object' && isHaloChatWebhook(body.request)) {
        body = { ...body.request, chatDetails: body.chatDetails };
    }

    // Agent joined: explicit event, Halo's webhook envelope, or a raw Halo chat message object (has chat_id, no messages array)
    if (body.event === 'agentJoined' || isHaloChatWebhook(body) || (body.chat_id && !body.chat && !Array.isArray(body.messages))) {
        try {
            await handleChatMessage(context, body);
        } catch (error) {
            context.log.error('Failed to handle chat message:', error.message);
            context.res = { status: 500, body: { error: error.message } };
        }
        return;
    }

    // Chat started: the raw Halo chat object, or { chat, ticket, agents } when the runbook enriches it
    const chat = body.chat || body;
    const ticket = body.ticket || null;
    const agents = body.agents;
    // Target channel: channelId in the body, then the CHAT_CHANNEL_ID app setting, then the stored (install) channel
    const channelId = (body.channelId || process.env.CHAT_CHANNEL_ID || '').trim() || null;

    if (!chat.id || !Array.isArray(chat.messages)) {
        // Log field names only (no values) so an unexpected Halo payload shape can be diagnosed
        context.log.warn(`Unrecognised chat payload, top-level fields: ${Object.keys(body).join(', ') || '(none)'}`);
        context.res = { status: 400, body: { error: 'Body must be a Halo chat object (id, messages) or { chat, ticket, agents, channelId }' } };
        return;
    }

    if (channelId && !isValidChannelId(channelId)) {
        context.res = { status: 400, body: { error: 'Invalid channelId: expected 19:xxxx@thread.tacv2 (URL-decoded from the channel link)' } };
        return;
    }

    const parsedEscalation = parseEscalation(body.escalation);
    const escalation = parsedEscalation && !parsedEscalation.warning ? parsedEscalation : null;
    if (parsedEscalation && parsedEscalation.warning) context.log.warn(`Chat ${chat.id}: ${parsedEscalation.warning}`);

    let payload = { chat, ticket, agents };
    const resolved = resolveTicket(chat, ticket);
    let mentionWarning = null;

    try {
        let sent;
        try {
            sent = await sendCard(cardFromPayload(payload, {}), channelId);
        } catch (error) {
            // Teams rejects the whole message if a tag mention is invalid (e.g. the tag is in another team).
            // Retry without tag mentions so the card still gets posted.
            const hasTags = Array.isArray(agents) && agents.some(a => a && a.type === 'tag');
            if (!hasTags || !isTagMentionError(error)) throw error;
            mentionWarning = `Sent without tag mentions: ${error.message}`;
            context.log.warn(`Chat ${chat.id}: send with tag mentions failed, retrying without them: ${error.message}`);
            payload = { ...payload, agents: agents.filter(a => !a || a.type !== 'tag') };
            sent = await sendCard(cardFromPayload(payload, {}), channelId);
        }
        if (!sent) {
            context.res = {
                status: 400,
                body: { error: 'Bot has not been installed in a team yet. No conversation reference stored.' }
            };
            return;
        }

        // "Chat so far" goes in the card's thread rather than on the card
        let transcriptActivityId = null;
        try {
            const transcriptCard = buildTranscriptCard(chat);
            if (transcriptCard) transcriptActivityId = await sendCardInThread(sent.conversationId, sent.activityId, transcriptCard);
        } catch (error) {
            context.log.warn(`Chat ${chat.id}: transcript reply not posted (the next message will retry): ${error.message}`);
        }

        // Remember where the card went so an agent joining can update it. The card is already sent,
        // so a failure here is logged rather than failing the request.
        const createdAt = new Date();
        let escalationDueAt = null;
        try {
            await saveChat(chat.id, {
                ...sent,
                transcriptActivityId,
                payload,
                escalation,
                lastMessageId: Number(chat.last_message_id || 0),
                createdAt: createdAt.toISOString()
            });
            // The escalate timer picks this up once it's due, unless an agent responds first
            if (escalation) {
                escalationDueAt = new Date(createdAt.getTime() + escalation.afterMinutes * 60000).toISOString();
                await addPending(chat.id, escalationDueAt);
            }
        } catch (error) {
            context.log.warn(`Card sent but chat ${chat.id} record or escalation not saved (agent join/escalation will not work for it): ${error.message}`);
        }

        context.log(`Chat notification sent for chat ${chat.id} (ticket ${resolved ? resolved.id : 'none'}, channel ${channelId || 'default'}${escalationDueAt ? `, escalation due ${escalationDueAt}` : ''})`);
        context.res = {
            status: 200,
            body: {
                success: true,
                chatId: chat.id,
                ticketId: resolved ? resolved.id : null,
                ...(escalationDueAt ? { escalationDueAt } : {}),
                ...(mentionWarning ? { warning: mentionWarning } : {}),
                ...(parsedEscalation && parsedEscalation.warning ? { escalationWarning: parsedEscalation.warning } : {})
            }
        };
    } catch (error) {
        context.log.error('Failed to send chat notification:', error.message);
        context.res = { status: 500, body: { error: error.message } };
    }
};
