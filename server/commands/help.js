const { MessageFactory } = require('botbuilder');

const GENERAL_COMMANDS = {
    'help': 'Show all available commands',
    'ima': 'Show your account role',
    'joke': 'Tell a random joke',
    'ping <host/ip>': 'Ping a hostname or IP address'
};

const ADMIN_COMMANDS = {
    '**Server Management**': null,
    'mark/set <server> available/unavailable': 'Set server availability flag',
    'mark/set <server> <field> <value>': 'Set a field on a server master record',
    'hyperv status': 'Show HyperV hosts currently processing',
    'processing status': 'Show processing queue status',
    'get global <var>': 'Read a global variable',
    'set global <var> <value>': 'Write a global variable',
    '**Lookups**': null,
    'where/find/search <ip>': 'Look up an IP in the asset database',
    'search asset <hostname/id>': 'Look up an asset by hostname or ID',
    '**Tickets**': null,
    'add <department> ticket <subject>': 'Create a support ticket (bugs, hardware, billing, etc.)',
    'add ticket #<id> post|reply <content>': 'Add a reply/post to an existing ticket',
    '**MailBaby**': null,
    'add mailbaby user <user> <pass>': 'Add a MailBaby SMTP user',
    'delete mailbaby user <user>': 'Delete a MailBaby SMTP user',
    'block help': 'Show blocked emails/domains commands',
    '**GitHub**': null,
    'github help': 'Show GitHub issues/labels commands'
};

module.exports = {
    match(text, lcText) {
        return lcText === 'help' ? {} : null;
    },
    async execute(_match, { context, ima }) {
        let text = '**Available Commands**\n\n';
        for (const [cmd, desc] of Object.entries(GENERAL_COMMANDS)) {
            text += `\`${ cmd }\` - ${ desc }\n`;
        }
        if (ima === 'admin') {
            text += '\n**Admin Commands**\n\n';
            for (const [cmd, desc] of Object.entries(ADMIN_COMMANDS)) {
                if (desc === null) {
                    text += `\n${ cmd }\n`;
                } else {
                    text += `\`${ cmd }\` - ${ desc }\n`;
                }
            }
        }
        await context.sendActivity(MessageFactory.text(text));
    }
};
