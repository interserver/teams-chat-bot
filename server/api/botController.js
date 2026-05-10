// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { BotActivityHandler } = require('../bot/botActivityHandler');
const { getAdapter } = require('../lib/adapter');
const { runWithRetry } = require('../lib/retry');

// botController's adapter is used for processActivity (inbound turn processing).
// It uses the shared adapter so credentials are pooled, but maintains its own
// onTurnError handler since turn processing has different error needs than
// proactive sends.
const adapter = getAdapter();

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
