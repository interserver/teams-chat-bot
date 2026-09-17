const { CardFactory, MessageFactory, TurnContext } = require('botbuilder');
const axios = require('axios');
const { bindCard, scheduleAutoDelete, AUTO_DELETE_MS } = require('../api/dailyRecapController');
const { MAX_ACTIVITY_BYTES, activityBytes, trimCard, fitTextActivity } = require('../lib/messageSize');

// Shared with the daily_recap_card_api.php endpoint on MyAdmin.
const RECAP_URL = process.env.DAILY_RECAP_URL
    || 'https://mystage.interserver.net/admin/ajax/daily_recap_card_api.php';
const RECAP_TOKEN = process.env.DAILY_RECAP_TOKEN || '';

// Best-effort plain-text digest used when the card itself cannot fit Teams'
// message size limit — keep the day's numbers visible, never lose the recap.
function recapFallbackText(data) {
    const stats = data && Array.isArray(data.stats) ? data.stats : [];
    const lines = stats
        .slice(0, 20)
        .map(s => `- ${ s.name || s.title || s.label || 'Service' }${ s.total != null ? `: ${ s.total }` : '' }`);
    const header = '⚠️ Daily recap card exceeded the Teams message size limit and was reduced to this summary.';
    return lines.length > 0 ? `${ header }\n\n${ lines.join('\n') }` : header;
}

module.exports = {
    match(text, lcText, { ima }) {
        // Exact match on the whole line, admin-only.  No partial match.
        return (ima === 'admin' && lcText === 'daily recap') ? {} : null;
    },
    async execute(_match, { context }) {
        if (!RECAP_TOKEN) {
            await context.sendActivity(MessageFactory.text(
                'Daily recap is unavailable — DAILY_RECAP_TOKEN is not set on the bot.'
            ));
            return;
        }
        let template; let data;
        try {
            const resp = await axios.get(RECAP_URL, {
                params:  { token: RECAP_TOKEN },
                headers: { 'X-Daily-Recap-Token': RECAP_TOKEN },
                timeout: 5 * 60 * 1000
            });
            template = resp.data && resp.data.template;
            data = resp.data && resp.data.data;
            if (!template || !data) {
                throw new Error('response missing template/data');
            }
        } catch (err) {
            const status = err.response && err.response.status;
            const detail = err.response && err.response.data ? JSON.stringify(err.response.data) : err.message;
            await context.sendActivity(MessageFactory.text(
                `Failed to fetch daily recap from MyAdmin${ status ? ' (HTTP ' + status + ')' : '' }: ${ detail }`
            ));
            return;
        }
        // Compact-mode trims so the bound card stays under Teams' 25 KB
        // Adaptive Card limit.  These are visibility flags consumed by
        // `$when` clauses in the template + a cleared pie URL so we
        // don't ship 11 KB of base64 the bot would never render.
        data.show_details  = false;
        data.show_pie      = false;
        data.show_month    = false;
        data.pie_chart_url = '';
        if (Array.isArray(data.stats)) {
            data.stats.forEach((s) => {
                s.orders       = [];
                s.empty_orders = true;
                s.order_count  = 0;
            });
        }

        let card;
        try {
            card = bindCard(template, data);
        } catch (err) {
            await context.sendActivity(MessageFactory.text(
                `Failed to bind daily recap template: ${ err.message }`
            ));
            return;
        }
        // Hard guard against Teams' ~25 KB serialized-activity limit: trim
        // the card, and if even a trimmed card cannot fit, send a short
        // plain-text summary instead of an oversized activity the SDK would
        // silently swallow (the RestError dies inside adapter.onTurnError).
        let outboundCard = card;
        let activity = { type: 'message', attachments: [CardFactory.adaptiveCard(outboundCard)] };
        if (activityBytes(activity) > MAX_ACTIVITY_BYTES) {
            outboundCard = trimCard(outboundCard);
            activity = { type: 'message', attachments: [CardFactory.adaptiveCard(outboundCard)] };
        }
        if (activityBytes(activity) > MAX_ACTIVITY_BYTES) {
            console.warn(`[dailyRecap] recap card still ${ activityBytes(activity) } bytes after trimming — sending text fallback`);
            await context.sendActivity(MessageFactory.text(fitTextActivity(recapFallbackText(data))));
            return;
        }
        const sent = await context.sendActivity(activity);
        // Auto-delete the recap after AUTO_DELETE_MS so it doesn't clutter
        // the channel.  Capture the conversation reference now while we're
        // still inside the turn — the timer fires after the turn ends.
        if (sent && sent.id) {
            const reference = TurnContext.getConversationReference(context.activity);
            scheduleAutoDelete(reference, sent.id, AUTO_DELETE_MS);
        }
    }
};
