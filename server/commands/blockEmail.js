const { MessageFactory } = require('botbuilder');

const EMAIL_RE = /(?<email>[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/;

module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        if (/^(blocks|block list|blocks list)$/i.test(lcText)) {
            return { action: 'list' };
        }
        if (new RegExp(`^(block|block email) ${ EMAIL_RE.source }$`, 'i').test(lcText)) {
            const m = lcText.match(new RegExp(EMAIL_RE, 'i'));
            return { action: 'add', email: m.groups.email };
        }
        if (new RegExp(`^(block remove|block delete|block email remove|block email delete|blocked email remove|blocked email delete) ${ EMAIL_RE.source }$`, 'i').test(lcText)) {
            const m = lcText.match(new RegExp(EMAIL_RE, 'i'));
            return { action: 'remove', email: m.groups.email };
        }
        return null;
    },
    async execute({ action, email }, { context, redis }) {
        if (action === 'list') {
            const blockedEmails = await redis.smembers('blocked_emails');
            blockedEmails.sort();
            const text = `*Blocked Emails* (${ blockedEmails.length })\n` + blockedEmails.join(', ');
            await context.sendActivity(MessageFactory.text(text));
        } else if (action === 'add') {
            const added = await redis.sadd('blocked_emails', email);
            if (added) {
                await context.sendActivity(MessageFactory.text(`✅ Successfully added *${ email }* to blocked emails list.`));
            } else {
                await context.sendActivity(MessageFactory.text(`⚠️ *${ email }* already exists in blocked emails list.`));
            }
        } else {
            const removed = await redis.srem('blocked_emails', email);
            if (removed) {
                await context.sendActivity(MessageFactory.text(`✅ Successfully removed *${ email }* from blocked emails list.`));
            } else {
                await context.sendActivity(MessageFactory.text(`⚠️ *${ email }* is not in blocked emails list.`));
            }
        }
    }
};
