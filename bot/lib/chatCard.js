const { validMentions, addMentions } = require('./teams');

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

// Halo chat timestamps are UTC without a zone. Adaptive Card DATE()/TIME() functions render
// them in each viewer's own time zone and locale, so UK and NZ agents both see local time.
function utcIso(ts) {
    if (!ts) return '';
    const date = new Date(/[zZ]|[+-]\d\d:\d\d$/.test(ts) ? ts : ts + 'Z');
    if (isNaN(date)) return '';
    return date.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

function localDateTime(ts) {
    const iso = utcIso(ts);
    return iso ? `{{DATE(${iso}, SHORT)}}, {{TIME(${iso})}}` : '';
}

function localTime(ts) {
    const iso = utcIso(ts);
    return iso ? `{{TIME(${iso})}}` : '';
}

function haloBaseUrl(chat) {
    if (process.env.HALO_URL) return process.env.HALO_URL.replace(/\/+$/, '');
    try {
        return new URL(chat.user_chat_url).origin;
    } catch {
        return null;
    }
}

// Halo variables that weren't set can come through as the text "null"/"undefined"; treat those as empty
function cleanValue(value) {
    if (value === null || value === undefined) return '';
    const text = String(value).trim();
    return /^(null|undefined)$/i.test(text) ? '' : text;
}

// Ticket numbers the end user typed in the chat flow (the "which ticket?" answer)
function typedTicketNumbers(chat) {
    return (chat.messages || [])
        .filter(isUserMessage)
        .map(m => (m.text || '').match(/^\s*#?(\d{1,10})\s*$/))
        .filter(Boolean)
        .map(match => match[1]);
}

function parseBool(value) {
    const text = cleanValue(value).toLowerCase();
    if (['true', 'yes', '1'].includes(text)) return true;
    if (['false', 'no', '0'].includes(text)) return false;
    return null;
}

// Work out which ticket the chat is about, in order of trust:
// 1. ticket.id supplied by the runbook (validated there), else ticket_id on the chat
// 2. the last end-user message that is just a number (typed in the flow)
// The ticket counts as existing if the user typed its number in the flow, otherwise it was raised
// from the chat. The runbook can override that with ticket.isNew.
function resolveTicket(chat, ticket) {
    const typed = typedTicketNumbers(chat);
    const runbookId = cleanValue(ticket && ticket.id);
    const id = /^\d+$/.test(runbookId) ? runbookId : (chat.ticket_id > 0 ? String(chat.ticket_id) : '');

    if (id) {
        const explicit = parseBool(ticket && ticket.isNew);
        const isNew = explicit !== null ? explicit : !typed.includes(id);
        return { id, source: isNew ? 'new ticket raised' : 'existing ticket' };
    }
    if (typed.length) return { id: typed[typed.length - 1], source: 'entered by user, not validated' };
    return null;
}

// Ticket reference as shown on the card, e.g. "VT:2148969" with TICKET_PREFIX "VT:" (default "#")
function ticketLabel(id) {
    return `${process.env.TICKET_PREFIX || '#'}${id}`;
}

// Keep the card well under Teams' message size limit on long chats
const MAX_TRANSCRIPT_ENTRIES = 20;

// Name of an agent in the chat, from the participants list or their messages
function agentNameFrom(chat, agentId) {
    if (!chat || !(Number(agentId) > 0)) return '';
    const participant = (chat.participants || []).find(p => Number(p.agent_id) === Number(agentId));
    if (participant && participant.name) return participant.name;
    const message = (chat.messages || []).filter(m => Number(m.agent_id) === Number(agentId)).pop();
    return (message && message.display_name) || '';
}

// A live chat message as a bubble: agents blue and indented from the right, the end user grey on
// the left, each labelled with their role and the time (in the viewer's own time zone)
function chatBubble(name, isAgent, timestamp, text) {
    const time = localTime(timestamp);
    const bubble = {
        type: 'Column',
        width: 85,
        items: [{
            type: 'Container',
            style: isAgent ? 'accent' : 'emphasis',
            items: [
                {
                    type: 'TextBlock',
                    text: `${name} (${isAgent ? 'Agent' : 'End user'})${time ? ` | ${time}` : ''}`,
                    size: 'Small',
                    weight: 'Bolder',
                    isSubtle: true,
                    wrap: true
                },
                { type: 'TextBlock', text, wrap: true, spacing: 'None' }
            ]
        }]
    };
    const spacer = { type: 'Column', width: 15, items: [] };
    return {
        type: 'ColumnSet',
        spacing: 'Small',
        columns: isAgent ? [spacer, bubble] : [bubble, spacer]
    };
}

// Halo message kinds, by system_id:
//   -1  chat flow (bot) message: a question, held until the user answers it
//    0  a real message from the end user (sender_id 1) or an agent (agent_id > 0)
//   >0  system event, e.g. "X joined the chat", "changed the current Ticket"
function buildTranscript(chat) {
    const entries = [];
    let pendingQuestion = '';

    for (const m of (chat && chat.messages) || []) {
        // Teams markdown ignores single newlines, so any run of line breaks becomes one paragraph break
        const text = stripHtml(m.text).replace(/\s*(\r?\n)+\s*/g, '\n\n');
        const systemId = Number(m.system_id);

        if (systemId === -1 || (m.sender_id === 0 && !(systemId > 0))) {
            if (text) pendingQuestion = text;
            continue;
        }
        if (!text) continue;

        if (systemId > 0) {
            entries.push([{ type: 'TextBlock', text: `_${truncate(text, 200)}_`, size: 'Small', isSubtle: true, wrap: true, spacing: 'Small' }]);
            continue;
        }

        const isAgent = Number(m.agent_id) > 0;
        if (!isAgent && pendingQuestion) {
            // Answer to a chat flow question
            entries.push([
                { type: 'TextBlock', text: truncate(pendingQuestion, 200), size: 'Small', isSubtle: true, wrap: true, spacing: 'Small' },
                { type: 'TextBlock', text: truncate(text, 500), weight: 'Bolder', wrap: true, spacing: 'None' }
            ]);
            pendingQuestion = '';
            continue;
        }

        // Live chat message
        const name = isAgent
            ? agentNameFrom(chat, m.agent_id) || m.display_name || 'Agent'
            : m.display_name || 'User';
        entries.push([chatBubble(name, isAgent, m.timestamp, truncate(text, 500))]);
        pendingQuestion = '';
    }

    const shown = entries.slice(-MAX_TRANSCRIPT_ENTRIES);
    const items = shown.flat();
    if (entries.length > shown.length) {
        items.unshift({
            type: 'TextBlock',
            text: `_${entries.length - shown.length} earlier messages not shown_`,
            size: 'Small',
            isSubtle: true,
            wrap: true,
            spacing: 'Small'
        });
    }
    return items;
}

// The card's state from a saved chat record: agent joined (green), escalated with no response (red),
// or waiting (blue). Pass {} for a card that has just been sent.
function chatStatus(record) {
    return {
        joined: !!record.joinedAt,
        agentName: record.joinedAgent || '',
        escalatedAfter: record.escalatedAt ? (record.escalation && record.escalation.afterMinutes) || '?' : null
    };
}

function buildCard(chat, ticket, resolved, baseUrl, hasAgents, status) {
    const participant = chat.other_participant || (chat.participants || [])[0] || {};
    const userName = participant.name || chat.display_name || 'Unknown user';
    const started = localDateTime((chat.messages || [])[0]?.timestamp);
    const ticketUrl = resolved && baseUrl ? `${baseUrl}/ticket?id=${resolved.id}` : null;

    let chatState = chat.livechat_started ? '' : ' | Waiting for an agent';
    let headerStyle = 'accent';
    if (status.joined) {
        chatState = status.agentName ? ` | Agent joined: ${status.agentName}` : ' | Agent joined';
        headerStyle = 'good';
    } else if (status.escalatedAfter) {
        chatState = ` | No response after ${status.escalatedAfter} min - escalated`;
        headerStyle = 'attention';
    }

    const facts = [];
    if (resolved) {
        facts.push({
            title: 'Ticket:',
            value: ticketUrl
                ? `[${ticketLabel(resolved.id)}](${ticketUrl}) (${resolved.source})`
                : `${ticketLabel(resolved.id)} (${resolved.source})`
        });
    } else {
        facts.push({ title: 'Ticket:', value: 'None linked' });
    }
    if (ticket) {
        const summary = stripHtml(cleanValue(ticket.summary));
        if (summary) facts.push({ title: 'Summary:', value: summary });
        if (cleanValue(ticket.status)) facts.push({ title: 'Status:', value: cleanValue(ticket.status) });
        if (cleanValue(ticket.priority)) facts.push({ title: 'Priority:', value: cleanValue(ticket.priority) });
        if (cleanValue(ticket.agent)) facts.push({ title: 'Agent:', value: cleanValue(ticket.agent) });
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

    const card = {
        type: 'AdaptiveCard',
        $schema: 'https://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.5',
        body: [
            {
                type: 'Container',
                style: headerStyle,
                items: [
                    {
                        type: 'TextBlock',
                        size: 'Medium',
                        weight: 'Bolder',
                        text: `New Live Chat - ${userName}${ticket && cleanValue(ticket.client) ? ` (${cleanValue(ticket.client)})` : ''}`,
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
            }
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

// "Chat so far" lives in a single reply in the card's thread, edited in place as the chat goes on.
// Edits don't notify thread followers, so the channel stays quiet. Returns null if there is nothing to show.
function buildTranscriptCard(chat) {
    const transcript = buildTranscript(chat);
    if (!transcript.length) return null;
    return {
        type: 'AdaptiveCard',
        $schema: 'https://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.5',
        body: [
            { type: 'TextBlock', text: 'Chat so far', weight: 'Bolder', wrap: true, id: 'transcript-heading' },
            ...transcript
        ],
        msteams: { width: 'Full' }
    };
}

// record: the saved chat record (or {} when first sending), for the card's joined/escalated state
function cardFromPayload(payload, record) {
    const { chat, ticket, agents } = payload;
    const card = buildCard(chat, ticket, resolveTicket(chat, ticket), haloBaseUrl(chat), validMentions(agents).length > 0, chatStatus(record || {}));
    addMentions(card, agents);
    return card;
}

const minutes = (n) => `${n} minute${Number(n) === 1 ? '' : 's'}`;

// Posted in the card's thread when no agent has responded within the threshold. A new reply
// (unlike an edit) notifies the people it @mentions.
function buildEscalationCard(afterMinutes, agents) {
    const card = {
        type: 'AdaptiveCard',
        $schema: 'https://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.5',
        body: [
            {
                type: 'Container',
                style: 'attention',
                items: [
                    {
                        type: 'TextBlock',
                        text: `No agent has responded to this chat after ${minutes(afterMinutes)}`,
                        weight: 'Bolder',
                        wrap: true
                    },
                    ...(validMentions(agents).length ? [{
                        type: 'TextBlock',
                        text: 'Escalating to',
                        size: 'Small',
                        isSubtle: true,
                        wrap: true,
                        spacing: 'Small',
                        id: 'agents-on-shift-label'
                    }] : [])
                ]
            }
        ],
        msteams: { width: 'Full' }
    };
    addMentions(card, agents);
    return card;
}

// Replaces the escalation reply once an agent responds, so the thread doesn't end on an unresolved
// red alert. It's an edit, so management aren't notified again (and there are no @mentions in it).
function buildEscalationResolvedCard(record) {
    const { afterMinutes, agents } = record.escalation || {};
    const responseMinutes = Math.max(1, Math.round((Date.parse(record.joinedAt) - Date.parse(record.createdAt)) / 60000));
    const who = record.joinedAgent || 'An agent';
    const names = (agents || []).map(a => a && a.displayName).filter(Boolean).join(', ');
    return {
        type: 'AdaptiveCard',
        $schema: 'https://adaptivecards.io/schemas/adaptive-card.json',
        version: '1.5',
        body: [
            {
                type: 'Container',
                style: 'good',
                items: [
                    {
                        type: 'TextBlock',
                        text: isNaN(responseMinutes) ? `${who} responded` : `${who} responded after ${minutes(responseMinutes)}`,
                        weight: 'Bolder',
                        wrap: true
                    },
                    {
                        type: 'TextBlock',
                        text: `Escalated after ${afterMinutes ? minutes(afterMinutes) : 'the threshold'} with no response${names ? ` (${names})` : ''}`,
                        size: 'Small',
                        isSubtle: true,
                        wrap: true,
                        spacing: 'Small'
                    }
                ]
            }
        ],
        msteams: { width: 'Full' }
    };
}

module.exports = {
    stripHtml, agentNameFrom, buildTranscriptCard, cardFromPayload,
    buildEscalationCard, buildEscalationResolvedCard, resolveTicket
};
