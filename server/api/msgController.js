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
const Redis = require('ioredis');

const TRANSIENT_RE = /ECONNRESET|ETIMEDOUT|ENOTFOUND|socket hang up/i;
const AUTH_ERROR_RE = /authorization has been denied|401|unauthorized/i;
const MAX_RETRIES = 2;
const RETRY_DELAY_MS = 1000;

function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }

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

async function sendProactiveMessage(conversationReference, messageText) {
    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
        try {
            await adapter.continueConversation(conversationReference, async (proactiveContext) => {
                await proactiveContext.sendActivity(messageText);
            });
            return; // success
        } catch (err) {
            const msg = err.message || '';
            const isTransient = TRANSIENT_RE.test(msg);
            const isAuth = AUTH_ERROR_RE.test(msg);

            if ((isTransient || isAuth) && attempt < MAX_RETRIES) {
                console.warn(`[proactive retry] Attempt ${attempt} failed (${msg}), retrying in ${RETRY_DELAY_MS}ms...`);

                if (isAuth) {
                    MicrosoftAppCredentials.trustServiceUrl(conversationReference.serviceUrl);
                }

                await sleep(RETRY_DELAY_MS * attempt);
                continue;
            }
            throw err;
        }
    }
}

// Handle errors during bot turn processing
adapter.onTurnError = async (context, error) => {
    const errorMsg = error.message || 'Oops. Something went wrong!';
    console.error(`\n [onTurnError] unhandled error: ${ error }`);

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
const redis = new Redis({ host: 'dragonfly.mailbaby.net', port: 6379 });
redis.on('connect', () => console.log("✅ Connected to Redis"));
redis.on('error', (err) => console.error("❌ Redis error:", err));

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
const msgHandler = async (req, res) => {
    const targetConversationId = '19:0c93975aae904b7db892891da3065c33@thread.v2';
    try {
        // retrieve stored reference
        const stored = await redis.get(`convref:${targetConversationId}`);
        if (stored) {
            const conversationReference = JSON.parse(stored);
            await sendProactiveMessage(conversationReference, req.body.message);
            res.json({ message: "sent" });
        } else {
            res.status(404).json({ message: "no conversation reference found" });
        }
    } catch (err) {
        console.error('[msgHandler] Error sending proactive message:', err.message);
        res.status(500).json({ message: err.message });
    }
};

module.exports = msgHandler;
