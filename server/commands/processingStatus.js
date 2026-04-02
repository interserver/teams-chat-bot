const { MessageFactory } = require('botbuilder');

module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        return /^processing status$/i.test(lcText) ? {} : null;
    },
    async execute(_match, { context, redis }) {
        try {
            const queueVal = await redis.get('processing_queue');
            const lastVal = await redis.get('processing_queue_last');
            const lastTime = lastVal ? new Date(parseInt(lastVal, 10) * 1000).toISOString().replace('T', ' ').slice(0, 19) : 'Never';
            let status;
            if (!queueVal || queueVal === '0') {
                status = 'Idle';
            } else {
                status = 'Processing';
            }
            const since = status === 'Processing' && queueVal
                ? new Date(parseInt(queueVal, 10) * 1000).toISOString().replace('T', ' ').slice(0, 19)
                : lastTime;
            await context.sendActivity(MessageFactory.text(`Processing Queue is **${ status }** since ${ since }`));
        } catch (err) {
            await context.sendActivity(MessageFactory.text(`Error checking processing status: ${ err.message }`));
            console.error('Processing status error:', err);
        }
    }
};
