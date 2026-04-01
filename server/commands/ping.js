const { MessageFactory } = require('botbuilder');

module.exports = {
    match(text, lcText) {
        const m = text.match(/^ping\s+(.+)$/i);
        return m ? { target: m[1].trim() } : null;
    },
    async execute({ target }, { context, bot, execFileAsync }) {
        if (bot.isValidHostname(target) || bot.isValidIP(target)) {
            await context.sendActivity(MessageFactory.text(`Pinging \`${ target }\` ...`));
            try {
                const args = ['-w', '10', '-W', '10', '-c', '4', '-q', target];
                const { stdout } = await execFileAsync('ping', args, { timeout: 15000 });
                const lines = stdout.trim().split('\n');
                const output = lines.slice(-3).join('\n');
                await context.sendActivity(MessageFactory.text('```\n' + output + '\n```'));
            } catch (err) {
                await context.sendActivity(MessageFactory.text(`⚠️ Error: ${ err.stderr || err.message }`));
            }
        } else {
            await context.sendActivity(MessageFactory.text(`❌ Invalid hostname or IP: \`${ target }\``));
        }
    }
};
