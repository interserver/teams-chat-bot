const { MessageFactory } = require('botbuilder');

module.exports = {
    match(text, lcText) {
        return lcText === 'ima' ? {} : null;
    },
    async execute(_match, { context, member, email, ima }) {
        await context.sendActivity(MessageFactory.text(
            `Hello ${ member.name }, I see your email is ${ email } and you are ima ${ ima }`
        ));
    }
};
