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
const { BotActivityHandler } = require('../bot/botActivityHandler');

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

    // Clear out state
    // await conversationState.delete(context);
    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${ error }`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity(errorMsg);

    // Uncomment the line below for local debugging
    await context.sendActivity(`Sorry, it looks like something went wrong. Exception Caught: ${ error }`);
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
        // Process bot activity
        await botActivityHandler.run(context);
    });
    // Route received a request to adapter for processing
    /*
    adapter.process(req, res, (context) => {
        // Process bot activity
        botActivityHandler.run(context);
    }); */
};

module.exports = botHandler;
