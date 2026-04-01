const { MessageFactory } = require('botbuilder');

const HOST_RE = /(?<host>([a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,})/;

module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        if (/^(blocked domains|blocked hosts|blocked domains list|blocked hosts list)$/i.test(lcText)) {
            return { action: 'list' };
        }
        if (new RegExp(`^(block domain|block hostname|block host) ${ HOST_RE.source }$`, 'i').test(lcText)) {
            const m = lcText.match(new RegExp(HOST_RE, 'i'));
            return { action: 'add', host: m.groups.host };
        }
        if (new RegExp(`^(block remove domain|block delete domain|block domain remove|block domain delete|blocked domain remove|blocked domain delete|blocked domains remove|blocked domains delete) ${ HOST_RE.source }$`, 'i').test(lcText)) {
            const m = lcText.match(new RegExp(HOST_RE, 'i'));
            return { action: 'remove', host: m.groups.host };
        }
        return null;
    },
    async execute({ action, host }, { context, redis }) {
        if (action === 'list') {
            const blockedDomains = await redis.smembers('blocked_domains');
            blockedDomains.sort();
            const text = `*Blocked Domains* (${ blockedDomains.length })\n` + blockedDomains.join(', ');
            await context.sendActivity(MessageFactory.text(text));
        } else if (action === 'add') {
            const added = await redis.sadd('blocked_domains', host);
            if (added) {
                await context.sendActivity(MessageFactory.text(`✅ Successfully added *${ host }* to blocked domains list.`));
            } else {
                await context.sendActivity(MessageFactory.text(`⚠️ *${ host }* already exists in blocked domains list.`));
            }
        } else {
            const removed = await redis.srem('blocked_domains', host);
            if (removed) {
                await context.sendActivity(MessageFactory.text(`✅ Successfully removed *${ host }* from blocked domains list.`));
            } else {
                await context.sendActivity(MessageFactory.text(`⚠️ *${ host }* is not in blocked domains list.`));
            }
        }
    }
};
