const { MessageFactory } = require('botbuilder');
const fs = require('fs');
const path = require('path');

module.exports = {
    match(text, lcText) {
        return (lcText === 'joke' || lcText === 'tell a joke') ? {} : null;
    },
    async execute(_match, { context }) {
        try {
            const jokesPath = path.join(__dirname, '../../jokes.json');
            const jokes = JSON.parse(fs.readFileSync(jokesPath, 'utf8'));
            const jokeList = Object.values(jokes).flat();
            if (Array.isArray(jokeList) && jokeList.length > 0) {
                const randomJoke = jokeList[Math.floor(Math.random() * jokeList.length)];
                if (Array.isArray(randomJoke)) {
                    for (const line of randomJoke) {
                        await context.sendActivity(MessageFactory.text(line));
                    }
                } else {
                    await context.sendActivity(MessageFactory.text(String(randomJoke)));
                }
            } else {
                await context.sendActivity(MessageFactory.text('Hmm, I don\'t have any jokes right now'));
            }
        } catch (err) {
            console.error('Error loading jokes.json:', err);
            await context.sendActivity(MessageFactory.text('⚠️ Sorry, I couldn\'t fetch a joke right now.'));
        }
    }
};
