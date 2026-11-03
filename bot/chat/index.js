const { isAuthorised, isValidChannelId, validMentions, addMentions, sendCard } = require('../lib/teams');

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

function buildCard(chat, ticket, resolved, baseUrl, hasAgents) {
    const participant = chat.other_participant || (chat.participants || [])[0] || {};
    const userName = participant.name || chat.display_name || 'Unknown user';
    const started = formatTimestamp((chat.messages || [])[0]?.timestamp);
    const ticketUrl = resolved && baseUrl ? `${baseUrl}/ticket?id=${resolved.id}` : null;

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
                style: 'accent',
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
                        text: `Chat #${chat.id}${chat.livechat_started ? '' : ' | Waiting for an agent'}`,
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

module.exports = async function (context, req) {
    context.log('Chat endpoint hit');

    if (!isAuthorised(req)) {
        context.res = { status: 401, body: { error: 'Unauthorized' } };
        return;
    }

    // Accept either the raw Halo chat object, or { chat, ticket, agents } when the runbook enriches it
    const body = req.body || {};
    const chat = body.chat || body;
    const ticket = body.ticket || null;
    const agents = body.agents;
    // Target channel: channelId in the body, then the CHAT_CHANNEL_ID app setting, then the stored (install) channel
    const channelId = (body.channelId || process.env.CHAT_CHANNEL_ID || '').trim() || null;

    if (!chat.id || !Array.isArray(chat.messages)) {
        context.res = { status: 400, body: { error: 'Body must be a Halo chat object (id, messages) or { chat, ticket, agents, channelId }' } };
        return;
    }

    if (channelId && !isValidChannelId(channelId)) {
        context.res = { status: 400, body: { error: 'Invalid channelId: expected 19:xxxx@thread.tacv2 (URL-decoded from the channel link)' } };
        return;
    }

    const resolved = resolveTicket(chat, ticket);
    const hasAgents = validMentions(agents).length > 0;
    const card = buildCard(chat, ticket, resolved, haloBaseUrl(chat), hasAgents);
    addMentions(card, agents);

    try {
        const sent = await sendCard(card, channelId);
        if (!sent) {
            context.res = {
                status: 400,
                body: { error: 'Bot has not been installed in a team yet. No conversation reference stored.' }
            };
            return;
        }

        context.log(`Chat notification sent for chat ${chat.id} (ticket ${resolved ? resolved.id : 'none'}, channel ${channelId || 'default'})`);
        context.res = { status: 200, body: { success: true, chatId: chat.id, ticketId: resolved ? resolved.id : null } };
    } catch (error) {
        context.log.error('Failed to send chat notification:', error.message);
        context.res = { status: 500, body: { error: error.message } };
    }
};
