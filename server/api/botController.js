// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    BotFrameworkAdapter
} = require('botbuilder');
const { BotActivityHandler } = require('../bot/botActivityHandler');
const { runWithRetry } = require('../lib/retry');

/*
const { TeamsBot } = require('../bot/teamsBot');
const { MainDialog } = require('../dialogs/mainDialog');

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(process.env);

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about how bots work.
// const adapter = new CloudAdapter(botFrameworkAuthentication);
*/
const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

// Handle errors during bot turn processing
adapter.onTurnError = async (context, error) => {
    const errorMsg = error.message || 'Oops. Something went wrong!';
    console.error(`\n [onTurnError] unhandled error: ${ error }`);

    // Don't attempt to send messages back if the error is a connection reset
    // or auth failure — those sends will also fail and cascade.
    const isTransient = /ECONNRESET|ETIMEDOUT|ENOTFOUND|socket hang up/i.test(errorMsg);
    const isAuthError = /authorization has been denied|401|unauthorized/i.test(errorMsg);

    if (isTransient) {
        console.error('[onTurnError] Transient network error, skipping reply to user.');
        return;
    }

    if (isAuthError) {
        console.error('[onTurnError] Authorization error — check MicrosoftAppId/MicrosoftAppPassword and bot registration.');
        return;
    }

    try {
        await context.sendTraceActivity(
            'OnTurnError Trace',
            `${ error }`,
            'https://www.botframework.com/schemas/error',
            'TurnError'
        );
        await context.sendActivity(`Sorry, it looks like something went wrong. Exception Caught: ${ errorMsg }`);
    } catch (sendError) {
        console.error(`[onTurnError] Failed to send error message to user: ${ sendError.message }`);
    }
};
/*
// Define the state store for your bot.
// See https://aka.ms/about-bot-state to learn more about using MemoryStorage.
const memoryStorage = new MemoryStorage();

// Create conversation and us er state with in-memory storage provider.
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);

// Create the main dialog.
const dialog = new MainDialog();
*/
// Create the bot that will handle incoming messages.
// const botActivityHandler = new TeamsBot(conversationState, userState, dialog);
const botActivityHandler = new BotActivityHandler();
const botHandler = (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        await runWithRetry(async () => {
            await botActivityHandler.run(context);
        }, {
            label: 'botController',
            serviceUrl: context.activity && context.activity.serviceUrl,
            maxRetries: 3
        });
    });
};

module.exports = botHandler;
module.exports.botActivityHandler = botActivityHandler;
