const { sendCardInThread, updateCardInThread, updateCard, isTagMentionError } = require('../lib/teams');
const { getChat, saveChat, listPending, removePending } = require('../lib/chatStore');
const { buildEscalationCard, buildEscalationResolvedCard, cardFromPayload } = require('../lib/chatCard');

// Pending escalations this far overdue (e.g. after an outage) are dropped rather than sent late
const STALE_AFTER_MS = 24 * 60 * 60 * 1000;

// Runs every minute: escalate chats where no agent has responded within the threshold the
// start runbook set (escalation.afterMinutes). Each chat is escalated at most once.
module.exports = async function (context) {
    const now = Date.now();
    const pending = await listPending();

    for (const { chatId, dueAt } of pending) {
        const due = Date.parse(dueAt);
        if (!isNaN(due) && due > now) continue;
        try {
            await escalate(context, chatId, isNaN(due) || now - due > STALE_AFTER_MS);
        } catch (error) {
            context.log.error(`Escalation for chat ${chatId} failed (will retry next run): ${error.message}`);
        }
    }
};

async function escalate(context, chatId, stale) {
    const found = await getChat(chatId);
    const record = found && found.record;

    // Nothing to do if an agent responded, the chat ended, or it was already escalated
    if (!record || !record.escalation || record.joinedAt || record.closedAt || record.escalatedAt || stale) {
        if (stale) context.log.warn(`Chat ${chatId}: escalation more than a day overdue, dropped`);
        await removePending(chatId);
        return;
    }

    // Claim it first (ETag) so an agent joining at the same moment, or an overlapping run, wins cleanly
    const claimedAt = new Date().toISOString();
    record.escalatedAt = claimedAt;
    if (!(await saveChat(chatId, record, found.etag))) return;

    const { afterMinutes, agents } = record.escalation;
    let replyId;
    try {
        replyId = await sendEscalation(context, chatId, record, afterMinutes, agents);
    } catch (error) {
        await releaseClaim(chatId, claimedAt);
        throw error;
    }
    await removePending(chatId);

    // Save the reply's ID so an agent responding later can turn it green. Use the record as saved
    // now, in case an agent joined while the escalation was being sent.
    const latest = (await rememberEscalationReply(chatId, replyId)) || record;

    try {
        if (latest.joinedAt && replyId) {
            await updateCardInThread(latest.conversationId, latest.activityId, replyId, buildEscalationResolvedCard(latest));
        }
        // Red header if still unanswered; green if an agent got there in the meantime
        await updateCard(latest.conversationId, latest.activityId, cardFromPayload(latest.payload, latest));
    } catch (error) {
        context.log.warn(`Chat ${chatId}: escalated, but card not updated: ${error.message}`);
    }

    context.log(`Chat ${chatId}: no agent response after ${afterMinutes} min, escalated`);
}

// Returns the updated record, or null if it couldn't be saved
async function rememberEscalationReply(chatId, replyId) {
    if (!replyId) return null;
    for (let attempt = 0; attempt < 3; attempt++) {
        const found = await getChat(chatId);
        if (!found) return null;
        found.record.escalationActivityId = replyId;
        if (await saveChat(chatId, found.record, found.etag)) return found.record;
    }
    return null;
}

// A new thread reply @mentioning management. As with the chat card, Teams rejects the whole
// message if a tag mention is invalid, so retry without tags rather than not escalating at all.
// Returns the reply's activity ID.
async function sendEscalation(context, chatId, record, afterMinutes, agents) {
    try {
        return await sendCardInThread(record.conversationId, record.activityId, buildEscalationCard(afterMinutes, agents));
    } catch (error) {
        const hasTags = agents.some(a => a && a.type === 'tag');
        if (!hasTags || !isTagMentionError(error)) throw error;
        context.log.warn(`Chat ${chatId}: escalation with tag mentions failed, retrying without them: ${error.message}`);
        return sendCardInThread(record.conversationId, record.activityId, buildEscalationCard(afterMinutes, agents.filter(a => !a || a.type !== 'tag')));
    }
}

// Undo the claim if the escalation couldn't be sent, so the next run tries again
async function releaseClaim(chatId, claimedAt) {
    const found = await getChat(chatId);
    if (!found || found.record.escalatedAt !== claimedAt) return;
    delete found.record.escalatedAt;
    await saveChat(chatId, found.record, found.etag);
}
