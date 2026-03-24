// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
/*    CloudAdapter,
    ConversationState,
    MemoryStorage,
    UserState,
    ConfigurationBotFrameworkAuthentication, */
    BotFrameworkAdapter
} = require('botbuilder');
const { MicrosoftAppCredentials } = require('botframework-connector');
const { BotActivityHandler } = require('../bot/botActivityHandler');

const TRANSIENT_RE = /ECONNRESET|ETIMEDOUT|ENOTFOUND|socket hang up/i;
const AUTH_ERROR_RE = /authorization has been denied|401|unauthorized/i;
const MAX_RETRIES = 2;
const RETRY_DELAY_MS = 1000;

function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }

async function runWithRetry(context, handler) {
    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
        try {
            await handler(context);
            return; // success
        } catch (err) {
            const msg = err.message || '';
            const isTransient = TRANSIENT_RE.test(msg);
            const isAuth = AUTH_ERROR_RE.test(msg);

            if ((isTransient || isAuth) && attempt < MAX_RETRIES) {
                console.warn(`[retry] Attempt ${attempt} failed (${msg}), retrying in ${RETRY_DELAY_MS}ms...`);

                // Force token refresh on auth errors
                if (isAuth) {
                    MicrosoftAppCredentials.trustServiceUrl(context.activity.serviceUrl);
                    adapter.credentials?.signRequest?.(null); // clear cached token
                }

                await sleep(RETRY_DELAY_MS * attempt);
                continue;
            }
            throw err; // final attempt or non-retryable error
        }
    }
}

async function sendProactiveMessage(conversationReference, messageText) {
    await adapter.continueConversation(conversationReference, async (proactiveContext) => {
        await proactiveContext.sendActivity(messageText);
    });
}

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
        await runWithRetry(context, async (ctx) => {
            await botActivityHandler.run(ctx);
        });
    });
};

module.exports = botHandler;
