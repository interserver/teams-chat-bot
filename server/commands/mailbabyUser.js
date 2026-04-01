const { MessageFactory } = require('botbuilder');

module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        let m;
        if ((m = lcText.match(/^add mailbaby user (\S+) (\S+)$/i))) {
            return { action: 'add', user: m[1], pass: m[2] };
        }
        if ((m = text.match(/^delete mailbaby user (\S+)$/i))) {
            return { action: 'delete', user: m[1] };
        }
        return null;
    },
    async execute({ action, user, pass }, { context, usersCollection }) {
        if (action === 'add') {
            const existing = await usersCollection.findOne({ username: user });
            if (existing) {
                await context.sendActivity(MessageFactory.text(`Found existing user '${ user }'`));
            } else {
                const result = await usersCollection.insertOne({ username: user, password: pass });
                if (result.insertedId) {
                    await context.sendActivity(MessageFactory.text(`Added user '${ user }' with password '${ pass }'`));
                } else {
                    await context.sendActivity(MessageFactory.text(`Error adding user '${ user }' with password '${ pass }'`));
                }
            }
        } else {
            const existing = await usersCollection.findOne({ username: user });
            if (existing) {
                await usersCollection.deleteOne({ username: user });
                await context.sendActivity(MessageFactory.text(`Removed user '${ user }'`));
            } else {
                await context.sendActivity(MessageFactory.text(`No user '${ user }' exists`));
            }
        }
    }
};
