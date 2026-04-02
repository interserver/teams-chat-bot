const { MessageFactory } = require('botbuilder');

module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        let m;
        if ((m = text.match(/^get global (\S+)$/i))) {
            return { action: 'get', varName: m[1] };
        }
        if ((m = text.match(/^set global (\S+) (.+)$/i))) {
            return { action: 'set', varName: m[1], value: m[2] };
        }
        return null;
    },
    async execute({ action, varName, value }, { context, redis }) {
        const key = `global:${ varName }`;
        try {
            if (action === 'get') {
                const val = await redis.get(key);
                if (val === null) {
                    await context.sendActivity(MessageFactory.text(`Global "${ varName }" is not set.`));
                } else {
                    await context.sendActivity(MessageFactory.text(`Global "${ varName }" value: ${ val }`));
                }
            } else {
                await redis.set(key, value);
                await context.sendActivity(MessageFactory.text(`Global "${ varName }" set to value: ${ value }`));
            }
        } catch (err) {
            await context.sendActivity(MessageFactory.text(`Error accessing global variable: ${ err.message }`));
            console.error('Global var error:', err);
        }
    }
};
