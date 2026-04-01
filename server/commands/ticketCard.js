const { CardFactory } = require('botbuilder');
const fs = require('fs');
const path = require('path');

module.exports = {
    match(text, lcText, { ima }) {
        return (ima === 'admin' && lcText === 'add ticket') ? {} : null;
    },
    async execute(_match, { context, bot }) {
        const cardPath = path.join(__dirname, '../../cards/add_ticket.json');
        const cardContents = JSON.parse(fs.readFileSync(cardPath, 'utf8'));
        const sentActivity = await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(cardContents)]
        });
        cardContents.body.forEach(element => bot.updateActionSubmitData(element, sentActivity));
        await context.updateActivity({
            type: 'message',
            id: sentActivity.id,
            conversation: context.activity.conversation,
            attachments: [CardFactory.adaptiveCard(cardContents)]
        });
    }
};
