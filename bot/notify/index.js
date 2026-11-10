const { isAuthorised, addMentions, sendCard } = require('../lib/teams');

module.exports = async function (context, req) {
    context.log('Notify endpoint hit');

    if (!isAuthorised(req)) {
        context.res = { status: 401, body: { error: 'Unauthorized' } };
        return;
    }

    const { agents, cardJson, ticketId, ticketSummary } = req.body;

    // Parse the adaptive card
    let card;
    try {
        card = typeof cardJson === 'string' ? JSON.parse(cardJson) : cardJson;
    } catch (e) {
        context.res = { status: 400, body: { error: 'Invalid cardJson: ' + e.message } };
        return;
    }

    // Convert HTML to markdown for Adaptive Card TextBlocks
    function htmlToMarkdown(html) {
        if (!html || typeof html !== 'string') return html;
        // Check if it actually contains HTML tags
        if (!/<[a-z][\s\S]*>/i.test(html)) return html;

        let text = html;
        // Line breaks
        text = text.replace(/<br\s*\/?>/gi, '\n');
        // Paragraphs
        text = text.replace(/<\/p>\s*<p[^>]*>/gi, '\n\n');
        text = text.replace(/<p[^>]*>/gi, '');
        text = text.replace(/<\/p>/gi, '\n\n');
        // Divs as line breaks
        text = text.replace(/<\/div>\s*<div[^>]*>/gi, '\n');
        text = text.replace(/<div[^>]*>/gi, '');
        text = text.replace(/<\/div>/gi, '\n');
        // Bold
        text = text.replace(/<(b|strong)[^>]*>([\s\S]*?)<\/\1>/gi, '**$2**');
        // Italic
        text = text.replace(/<(i|em)[^>]*>([\s\S]*?)<\/\1>/gi, '_$2_');
        // Links
        text = text.replace(/<a[^>]+href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, '[$2]($1)');
        // Unordered lists
        text = text.replace(/<ul[^>]*>([\s\S]*?)<\/ul>/gi, (_, inner) => {
            return inner.replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, '- $1\n');
        });
        // Ordered lists
        text = text.replace(/<ol[^>]*>([\s\S]*?)<\/ol>/gi, (_, inner) => {
            let n = 0;
            return inner.replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, (__, content) => {
                n++;
                return `${n}. ${content}\n`;
            });
        });
        // Headings
        text = text.replace(/<h1[^>]*>([\s\S]*?)<\/h1>/gi, '**$1**\n');
        text = text.replace(/<h2[^>]*>([\s\S]*?)<\/h2>/gi, '**$1**\n');
        text = text.replace(/<h3[^>]*>([\s\S]*?)<\/h3>/gi, '**$1**\n');
        // Horizontal rule
        text = text.replace(/<hr[^>]*\/?>/gi, '\n---\n');
        // Strip remaining HTML tags
        text = text.replace(/<[^>]+>/g, '');
        // Decode common HTML entities
        text = text.replace(/&amp;/g, '&');
        text = text.replace(/&lt;/g, '<');
        text = text.replace(/&gt;/g, '>');
        text = text.replace(/&quot;/g, '"');
        text = text.replace(/&#39;/g, "'");
        text = text.replace(/&nbsp;/g, ' ');
        // Clean up excessive newlines
        text = text.replace(/\n{3,}/g, '\n\n');
        return text.trim();
    }

    // Evaluate $when conditions and remove non-matching elements
    // Supports: == (equals), != (not equals)
    // Values are compared case-insensitively
    // Truthy values: 'yes', 'true', '1' are treated as equivalent
    function evaluateWhen(expr) {
        // Match: left == 'right' or left != 'right'
        const match = expr.match(/^(.+?)\s*(==|!=)\s*'(.+?)'$/);
        if (!match) return true; // Unknown expression format — keep the element

        let left = match[1].trim();
        const operator = match[2];
        let right = match[3];

        // Normalise truthy/falsy values for boolean comparisons
        const truthyValues = ['yes', 'true', '1'];
        const falsyValues = ['no', 'false', '0'];
        const leftLower = left.toLowerCase();
        const rightLower = right.toLowerCase();

        if (truthyValues.includes(leftLower)) left = 'true';
        else if (falsyValues.includes(leftLower)) left = 'false';
        else left = leftLower;

        if (truthyValues.includes(rightLower)) right = 'true';
        else if (falsyValues.includes(rightLower)) right = 'false';
        else right = rightLower;

        if (operator === '==') return left === right;
        if (operator === '!=') return left !== right;
        return true;
    }

    function processWhen(obj) {
        if (Array.isArray(obj)) {
            return obj.filter(item => {
                if (item && item.$when) {
                    if (!evaluateWhen(item.$when)) return false;
                    delete item.$when;
                }
                // Recurse into child arrays
                for (const key of Object.keys(item || {})) {
                    if (Array.isArray(item[key])) {
                        item[key] = processWhen(item[key]);
                    }
                }
                return true;
            });
        }
        return obj;
    }

    if (card.body) {
        card.body = processWhen(card.body);
    }

    // Convert HTML in TextBlock text fields to markdown
    function processHtml(elements) {
        if (!Array.isArray(elements)) return;
        for (const el of elements) {
            if (el.type === 'TextBlock' && el.text) {
                el.text = htmlToMarkdown(el.text);
            }
            if (el.items) processHtml(el.items);
            if (el.columns) {
                for (const col of el.columns) {
                    if (col.items) processHtml(col.items);
                }
            }
        }
    }
    processHtml(card.body);

    // Use full width for the card in Teams
    card.msteams = { ...(card.msteams || {}), width: 'Full' };

    addMentions(card, agents);

    // Send the proactive message
    try {
        // NOTIFY_CHANNEL_ID pins OOH alerts to a channel, so installing the bot in another team
        // (which replaces the stored conversation reference) doesn't move them
        const sent = await sendCard(card, (process.env.NOTIFY_CHANNEL_ID || '').trim() || null);
        if (!sent) {
            context.res = {
                status: 400,
                body: { error: 'Bot has not been installed in a team yet. No conversation reference stored.' }
            };
            return;
        }

        context.log(`Notification sent for ticket ${ticketId}`);
        context.res = { status: 200, body: { success: true, ticketId } };
    } catch (error) {
        context.log.error('Failed to send proactive message:', error.message);
        context.res = { status: 500, body: { error: error.message } };
    }
};
