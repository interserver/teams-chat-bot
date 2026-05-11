const { CardFactory, MessageFactory, TurnContext } = require('botbuilder');
const axios = require('axios');
const { bindCard, scheduleAutoDelete, AUTO_DELETE_MS } = require('../api/dailyRecapController');

// Shared with the daily_recap_card_api.php endpoint on MyAdmin.
const RECAP_URL = process.env.DAILY_RECAP_URL
    || 'https://mystage.interserver.net/admin/ajax/daily_recap_card_api.php';
const RECAP_TOKEN = process.env.DAILY_RECAP_TOKEN || '';

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
        const cardSize = JSON.stringify(card).length;
        if (cardSize > 25 * 1024) {
            console.warn(`[dailyRecap] bound card is ${ Math.round(cardSize / 1024) } KB — Teams' 25 KB limit may drop it`);
        }
        const sent = await context.sendActivity({
            type: 'message',
            attachments: [CardFactory.adaptiveCard(card)]
        });
        // Auto-delete the recap after AUTO_DELETE_MS so it doesn't clutter
        // the channel.  Capture the conversation reference now while we're
        // still inside the turn — the timer fires after the turn ends.
        if (sent && sent.id) {
            const reference = TurnContext.getConversationReference(context.activity);
            scheduleAutoDelete(reference, sent.id, AUTO_DELETE_MS);
        }
    }
};
