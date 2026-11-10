const { isAuthorised, isValidChannelId, validMentions, addMentions, sendCard, updateCard, replyInThread } = require('../lib/teams');
const { getChat, saveChat } = require('../lib/chatStore');

// Halo chat messages from the bot/flow have sender_id 0; anything else is the end-user
const isUserMessage = (m) => m && m.sender_id !== 0;

function stripHtml(text) {
    if (!text || typeof text !== 'string') return '';
    return text
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&nbsp;/g, ' ')
        .trim();
}

function truncate(text, max) {
    return text.length > max ? text.slice(0, max - 3) + '...' : text;
}

// Halo returns timestamps without a zone; treat them as UTC and show UK time
function formatTimestamp(ts) {
    if (!ts) return '';
    const date = new Date(/[zZ]|[+-]\d\d:\d\d$/.test(ts) ? ts : ts + 'Z');
    if (isNaN(date)) return '';
    return date.toLocaleString('en-GB', {
        timeZone: 'Europe/London',
        day: '2-digit', month: 'short', year: 'numeric',
        hour: '2-digit', minute: '2-digit'
    });
}

function haloBaseUrl(chat) {
    if (process.env.HALO_URL) return process.env.HALO_URL.replace(/\/+$/, '');
    try {
        return new URL(chat.user_chat_url).origin;
    } catch {
        return null;
    }
}

// Work out which ticket the chat is about, in order of trust:
// 1. ticket.id supplied by the runbook (validated there)
// 2. ticket_id on the chat (ticket logged from the chat)
// 3. the last end-user message that is just a number (typed in the flow)
function resolveTicket(chat, ticket) {
    if (ticket && ticket.id) return { id: String(ticket.id), source: 'existing ticket' };
    if (chat.ticket_id > 0) return { id: String(chat.ticket_id), source: 'logged from chat' };

    const typed = (chat.messages || [])
        .filter(isUserMessage)
        .map(m => (m.text || '').match(/^\s*#?(\d{1,10})\s*$/))
        .filter(Boolean)
        .pop();
    if (typed) return { id: typed[1], source: 'entered by user, not validated' };

    return null;
}

// Pair each end-user answer with the flow question that preceded it
function buildTranscript(messages) {
    const items = [];
    let lastQuestion = '';
    for (const m of messages || []) {
        if (!isUserMessage(m)) {
            lastQuestion = stripHtml(m.text);
            continue;
        }
        items.push(
            {
                type: 'TextBlock',
                text: truncate(lastQuestion || 'Message', 200),
                size: 'Small',
                isSubtle: true,
                wrap: true,
                spacing: 'Small'
            },
            {
                type: 'TextBlock',
                text: truncate(stripHtml(m.text), 500),
                weight: 'Bolder',
                wrap: true,
                spacing: 'None'
            }
        );
        lastQuestion = '';
    }
    return items;
}

// joined: { agentName } once an agent has joined the chat, otherwise null
function buildCard(chat, ticket, resolved, baseUrl, hasAgents, joined) {
    const participant = chat.other_participant || (chat.participants || [])[0] || {};
    const userName = participant.name || chat.display_name || 'Unknown user';
    const started = formatTimestamp((chat.messages || [])[0]?.timestamp);
    const ticketUrl = resolved && baseUrl ? `${baseUrl}/ticket?id=${resolved.id}` : null;

    let chatState = chat.livechat_started ? '' : ' | Waiting for an agent';
    if (joined) chatState = joined.agentName ? ` | Agent joined: ${joined.agentName}` : ' | Agent joined';

    const facts = [];
    if (resolved) {
        facts.push({
            title: 'Ticket:',
            value: ticketUrl ? `[#${resolved.id}](${ticketUrl}) (${resolved.source})` : `#${resolved.id} (${resolved.source})`
        });
    } else {
        facts.push({ title: 'Ticket:', value: 'None linked' });
    }
    if (ticket) {
        if (ticket.summary) facts.push({ title: 'Summary:', value: stripHtml(ticket.summary) });
        if (ticket.status) facts.push({ title: 'Status:', value: ticket.status });
        if (ticket.priority) facts.push({ title: 'Priority:', value: ticket.priority });
        if (ticket.agent) facts.push({ title: 'Agent:', value: ticket.agent });
    }

    const details = [{ type: 'FactSet', facts }];
    if (hasAgents) {
        details.push({
            type: 'TextBlock',
            text: 'Notifying',
            size: 'Small',
            isSubtle: true,
            wrap: true,
            spacing: 'Medium',
            id: 'agents-on-shift-label'
        });
    }

    // Transcript gets its own box below the details, so it reads as a separate section
    const transcript = buildTranscript(chat.messages);
    const transcriptContainer = transcript.length ? [{
        type: 'Container',
        spacing: 'Medium',
        style: 'emphasis',
        showBorder: true,
        roundedCorners: true,
        id: 'transcript-container',
        items: [
            {
                type: 'TextBlock',
                text: 'Chat so far',
                weight: 'Bolder',
                wrap: true
            },
            ...transcript
        ]
    }] : [];

    const card = {
        type: 'AdaptiveCard',
        $schema: 'https://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.5',
        body: [
            {
                type: 'Container',
                style: joined ? 'good' : 'accent',
                items: [
                    {
                        type: 'TextBlock',
                        size: 'Medium',
                        weight: 'Bolder',
                        text: `New Live Chat - ${userName}${ticket && ticket.client ? ` (${ticket.client})` : ''}`,
                        style: 'heading',
                        wrap: true
                    },
                    {
                        type: 'TextBlock',
                        text: `Chat #${chat.id}${chatState}`,
                        isSubtle: true,
                        spacing: 'None',
                        wrap: true
                    },
                    ...(started ? [{
                        type: 'TextBlock',
                        text: `Started ${started}`,
                        isSubtle: true,
                        spacing: 'None',
                        wrap: true
                    }] : [])
                ]
            },
            {
                type: 'Container',
                spacing: 'Small',
                style: 'emphasis',
                id: 'details-container',
                items: details
            },
            ...transcriptContainer
        ],
        actions: [],
        msteams: { width: 'Full' },
        speak: `New live chat from ${userName}`
    };

    // "Open Halo" only as a fallback when there is no ticket to link to
    if (ticketUrl) card.actions.push({ type: 'Action.OpenUrl', title: 'View Ticket in Halo', url: ticketUrl });
    else if (baseUrl) card.actions.push({ type: 'Action.OpenUrl', title: 'Open Halo', url: baseUrl });

    return card;
}

function cardFromPayload(payload, joined) {
    const { chat, ticket, agents } = payload;
    const card = buildCard(chat, ticket, resolveTicket(chat, ticket), haloBaseUrl(chat), validMentions(agents).length > 0, joined);
    addMentions(card, agents);
    return card;
}

// Agent joined: update the original card, or reply in its thread if the update fails.
// Halo's webhook fires on every agent message, so only the first one for a chat does anything.
// Calls that don't apply return 200 with "ignored" so Halo doesn't treat them as webhook failures.
async function handleAgentJoined(context, body) {
    const chatId = body.chatId || body.chat_id;
    const agentName = String(body.agentName || body.agent_name || body.display_name || '').trim();

    if (!chatId) {
        context.res = { status: 400, body: { error: 'agentJoined needs chatId' } };
        return;
    }

    // A raw chat message from the end user or the flow (agent_id 0) is not an agent joining
    if (body.agent_id !== undefined && !(Number(body.agent_id) > 0)) {
        context.res = { status: 200, body: { success: true, chatId, ignored: 'not an agent message' } };
        return;
    }

    const found = await getChat(chatId);
    if (!found) {
        context.log(`Agent message for chat ${chatId} ignored: no card was sent for it`);
        context.res = { status: 200, body: { success: true, chatId, ignored: 'no card sent for this chat' } };
        return;
    }

    const { record, etag } = found;
    if (record.joinedAt) {
        context.res = { status: 200, body: { success: true, chatId, alreadyUpdated: true } };
        return;
    }

    // Claim the update first so two messages arriving together only update the card once
    record.joinedAt = new Date().toISOString();
    record.joinedAgent = agentName;
    if (!(await saveChat(chatId, record, etag))) {
        context.res = { status: 200, body: { success: true, chatId, alreadyUpdated: true } };
        return;
    }

    let method = 'updated';
    try {
        await updateCard(record.conversationId, record.activityId, cardFromPayload(record.payload, { agentName }));
    } catch (error) {
        context.log.warn(`Card update failed for chat ${chatId}, replying in thread instead: ${error.message}`);
        method = 'replied';
        await replyInThread(record.conversationId, record.activityId, `${agentName || 'An agent'} has joined the chat.`);
    }

    context.log(`Agent joined chat ${chatId} (${agentName || 'name not supplied'}), card ${method}`);
    context.res = { status: 200, body: { success: true, chatId, method } };
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

    // Agent joined: explicit event, or a raw Halo chat message object (has chat_id, no messages array)
    if (body.event === 'agentJoined' || (body.chat_id && !body.chat && !Array.isArray(body.messages))) {
        try {
            await handleAgentJoined(context, body);
        } catch (error) {
            context.log.error('Failed to handle agent joined:', error.message);
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

    const payload = { chat, ticket, agents };
    const resolved = resolveTicket(chat, ticket);
    const card = cardFromPayload(payload, null);

    try {
        const sent = await sendCard(card, channelId);
        if (!sent) {
            context.res = {
                status: 400,
                body: { error: 'Bot has not been installed in a team yet. No conversation reference stored.' }
            };
            return;
        }

        // Remember where the card went so an agent joining can update it. The card is already sent,
        // so a failure here is logged rather than failing the request.
        try {
            await saveChat(chat.id, { ...sent, payload, createdAt: new Date().toISOString() });
        } catch (error) {
            context.log.warn(`Card sent but chat ${chat.id} record not saved (agent join will not update it): ${error.message}`);
        }

        context.log(`Chat notification sent for chat ${chat.id} (ticket ${resolved ? resolved.id : 'none'}, channel ${channelId || 'default'})`);
        context.res = { status: 200, body: { success: true, chatId: chat.id, ticketId: resolved ? resolved.id : null } };
    } catch (error) {
        context.log.error('Failed to send chat notification:', error.message);
        context.res = { status: 500, body: { error: error.message } };
    }
};
