const { MessageFactory } = require('botbuilder');
const axios = require('axios');

const TICKET_URL = process.env.TICKET_API_URL || 'https://mystage.interserver.net/admin/ajax/create_ticket.php';

module.exports = {
    match(text, lcText, { ima, context }) {
        if (ima !== 'admin') return null;
        const val = context.activity.value;
        if (val && val.msteams && val.msteams.type === 'addTicketCancel') {
            return { action: 'cancel' };
        }
        if (val && val.msteams && val.msteams.type === 'addTicketSubmit') {
            return { action: 'submit' };
        }
        return null;
    },
    async execute({ action }, { context, member, email }) {
        if (action === 'cancel') {
            console.log(context.activity.value);
            await context.updateActivity({
                type: 'message',
                id: context.activity.value.activityId,
                conversation: context.activity.conversation,
                text: 'Add ticket canceled'
            });
            return;
        }
        // action === 'submit'
        console.log(context.activity.value);
        try {
            const { subject, contents: body, department: dept, priority, status, type, notes } = context.activity.value;
            const name = member.name;
            const params = { subject, body, dept, email, name, priority, status, type };
            if (notes) {
                params.notes = notes;
            }
            const response = await axios.post(TICKET_URL,
                new URLSearchParams(params),
                { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } });
            if (response.status === 200) {
                await context.updateActivity({
                    type: 'message',
                    id: context.activity.value.activityId,
                    conversation: context.activity.conversation,
                    text: response.data
                });
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
