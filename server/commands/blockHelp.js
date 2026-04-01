const { MessageFactory } = require('botbuilder');

module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        return /^blocks? help$/i.test(lcText) ? {} : null;
    },
    async execute(_match, { context }) {
        const commands = {
            'blocks list': { description: 'List all blocked Emails' },
            'block <email>': { description: 'Adds an email to the blocked emails list.' },
            'block remove <email>': { description: 'Removes an email address from the blocked emails list.' },
            'blocked domains': { description: 'List all blocked Domains' },
            'block domain <host>': { description: 'Adds a domain to the blocked domains list.' },
            'block domain remove <host>': { description: 'Removes a domain from the blocked domains list.' },
            'block help': { description: 'Show all available Blocked Email/Domains commands' }
        };
        let text = '*MailBaby Blocked Emails and Domains Help*\n';
        for (const [command, details] of Object.entries(commands)) {
            text += `\`${ command }\` - ${ details.description }\n`;
        }
        await context.sendActivity(MessageFactory.text(text));
    }
};
