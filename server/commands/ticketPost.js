const { MessageFactory } = require('botbuilder');
const axios = require('axios');

const TICKET_POST_URL = process.env.TICKET_POST_API_URL || 'https://mystage.interserver.net/admin/ajax/post_ticket.php';

// Ticket id can be a plain integer (e.g. 12345) or a mask id like #EBC-923-68152
// (alphanumeric segments joined by dashes, optional leading #).
const TICKET_POST_RE = /^add ticket #?(\d+|[A-Za-z0-9]+(?:-[A-Za-z0-9]+)+) (?:post|reply) ([\s\S]+)$/msi;

module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        const m = text.match(TICKET_POST_RE);
        if (!m) return null;
        return { ticketId: m[1], body: m[2].trim() };
    },
    async execute({ ticketId, body }, { context, member, email }) {
        const name = member.name;
        try {
            const response = await axios.post(TICKET_POST_URL,
                new URLSearchParams({ ticket_id: ticketId, body, email, name }),
                { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } });
            if (response.status === 200) {
                await context.sendActivity(MessageFactory.text(response.data));
                console.log('✅ Success:', response.data);
            } else {
                await context.sendActivity(MessageFactory.text(response.data));
                console.log(`⚠️ Unexpected status ${ response.status }:`, response.data);
            }
        } catch (error) {
            if (error.response) {
                await context.sendActivity(MessageFactory.text(error.response.data));
                console.log(`❌ Error ${ error.response.status }:`, error.response.data);
            } else {
                await context.sendActivity(MessageFactory.text(error.message));
                console.log('❌ Request failed:', error.message);
            }
        }
    }
};
