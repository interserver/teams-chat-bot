---
page_type: sample
description: This bot can capture all channel messages in Teams using RSC permissions, without the need for @mentions.
products:
- office-teams
- office
- office-365
languages:
- nodejs
extensions:
 contentType: samples
 createdDate: "06/10/2021 01:48:56 AM"
urlFragment: officedev-microsoft-teams-samples-bot-receive-channel-messages-withRSC-nodejs
---

# Receive Channel messages with RSC permissions

This sample app illustrates how a bot can capture all channel messages in Microsoft Teams by utilizing RSC (resource-specific consent) permissions, eliminating the need for @mentions. The bot supports adaptive card responses, easy local testing with tools like ngrok or dev tunnels, and deployment to Azure, allowing it to function effectively across different channels and group chats in Teams.

This feature shown in this sample is currently available in Public Developer Preview only.

## Included Features
* Bots
* Adaptive Cards
* RSC Permissions

## Interaction with app

![Bot Receive Channel MessagesWithRSCGif](images/Bot_Channel_Messenging-RSC-nodejs-gif.gif)

## Try it yourself - experience the App in your Microsoft Teams client
Please find below demo manifest which is deployed on Microsoft Azure and you can try it yourself by uploading the app package (.zip file link below) to your teams and/or as a personal app. (Uploading must be enabled for your tenant, [see steps here](https://docs.microsoft.com/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant#enable-custom-teams-apps-and-turn-on-custom-app-uploading)).

**Receive Channel messages with RSC permissions:** [Manifest](/samples/bot-receive-channel-messages-withRSC/csharp/demo-manifest/Bot-RSC.zip)

## Prerequisites

1. Office 365 tenant. You can get a free tenant for development use by signing up for the [Office 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program).

2. To test locally, [NodeJS](https://nodejs.org/en/download/) must be installed on your development machine (version 16.14.2  or higher).

    ```bash
    # determine node version
    node --version
    ```
3. [dev tunnel](https://learn.microsoft.com/en-us/azure/developer/dev-tunnels/get-started?tabs=windows) or [Ngrok](https://ngrok.com/download) (For local environment testing) latest version (any other tunneling software can also be used)

   If you are using Ngrok to test locally, you'll need [Ngrok](https://ngrok.com/) installed on your development machine.
   Make sure you've downloaded and installed Ngrok on your local machine. ngrok will tunnel requests from the Internet to your local computer and terminate the SSL connection from Teams.

4. [Microsoft 365 Agents Toolkit for VS Code](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension) or [TeamsFx CLI](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-cli?pivots=version-one)

## Run the app (Using Microsoft 365 Agents Toolkit for Visual Studio Code)

The simplest way to run this sample in Teams is to use Microsoft 365 Agents Toolkit for Visual Studio Code.

1. Ensure you have downloaded and installed [Visual Studio Code](https://code.visualstudio.com/docs/setup/setup-overview)
1. Install the [Microsoft 365 Agents Toolkit extension](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension)
1. Select **File > Open Folder** in VS Code and choose this samples directory from the repo
1. Using the extension, sign in with your Microsoft 365 account where you have permissions to upload custom apps
1. Select **Debug > Start Debugging** or **F5** to run the app in a Teams web client.
1. In the browser that launches, select the **Add** button to install the app to Teams.

> If you do not have permission to upload custom apps (uploading), Microsoft 365 Agents Toolkit will recommend creating and using a Microsoft 365 Developer Program account - a free program to get your own dev environment sandbox that includes Teams.

## Setup

> NOTE: The free ngrok plan will generate a new URL every time you run it, which requires you to update your Azure AD registration, the Teams app manifest, and the project configuration. A paid account with a permanent ngrok URL is recommended.

1) Setup for Bot
- Register Azure AD application
- Register a bot with Azure Bot Service, following the instructions [here](https://docs.microsoft.com/azure/bot-service/bot-service-quickstart-registration?view=azure-bot-service-3.0).

- Ensure that you've [enabled the Teams Channel](https://docs.microsoft.com/azure/bot-service/channel-connect-teams?view=azure-bot-service-4.0)
- While registering the bot, use `https://<your_tunnel_domain>/api/messages` as the messaging endpoint.
    
    > NOTE: When you create your app registration in Azure portal, you will create an App ID and App password - make sure you keep these for later.

2) Setup NGROK
 - Run ngrok - point to port 3978

   ```bash
   ngrok http 3978 --host-header="localhost:3978"
   ```  

   Alternatively, you can also use the `dev tunnels`. Please follow [Create and host a dev tunnel](https://learn.microsoft.com/en-us/azure/developer/dev-tunnels/get-started?tabs=windows) and host the tunnel with anonymous user access command as shown below:

   ```bash
   devtunnel host -p 3978 --allow-anonymous
   ```

3) Setup for code
- Clone the repository

    ```bash
    git clone https://github.com/OfficeDev/Microsoft-Teams-Samples.git
    ```

- In the folder where repository is cloned navigate to `samples/bot-receive-channel-messages-withRSC/nodejs`

- Install node modules

   Inside node js folder, open your local terminal and run the below command to install node modules. You can do the same in Visual Studio code terminal by opening the project in Visual Studio code.

    ```bash
    npm install
    ```
- Update the `.env` configuration for the bot to use the `MicrosoftAppId` (Microsoft App Id) and `MicrosoftAppPassword` (App Password) from the Microsoft Entra ID app registration in Azure portal or from Bot Framework registration. 
> NOTE: the App Password is referred to as the `client secret` in the azure portal app registration service and you can always create a new client secret anytime.

- Run your app
    ```bash
    npm start
    ```
- Install modules & Run the NodeJS Server
  - Server will run on PORT: 3978
  - Open a terminal and navigate to project root directory
  
  ```bash
    npm run server
  ```
> NOTE:This command is equivalent to: npm install > npm start

4) Run your app

    ```bash
    npm start
    ```
5) Setup Manifest for Teams

    - **Edit** the `manifest.json` contained in the `appManifest` folder to replace your Microsoft App Id (that was created when you registered your bot earlier) *everywhere* you see the place holder string `<<YOUR-MICROSOFT-APP-ID>>` (depending on the scenario the Microsoft App Id may occur multiple times in the `manifest.json`) 
        `<<DOMAIN-NAME>>` with base Url domain. E.g. if you are using ngrok it would be `https://1234.ngrok-free.app` then your domain-name will be `1234.ngrok-free.app` and if you are using dev tunnels then your domain will be like: `12345.devtunnels.ms`.
         Replace <<MANIFEST-ID>> with any GUID or with your MicrosoftAppId/app id

    - **Zip** up the contents of the `appManifest` folder to create a `manifest.zip`
    - **Upload** in a team to test
         - Select or create a team
         - Select the ellipses **...** from the left pane. The drop-down menu appears.
         - Select **Manage Team**, then select **Apps** 
         - Then select **Upload a custom app** from the lower right corner.
         - Then select the `manifest.zip` file from `appManifest`, and then select **Add** to add the bot to your selected team.

**Note**: If you are facing any issue in your app, please uncomment [this](https://github.com/OfficeDev/Microsoft-Teams-Samples/blob/main/samples/bot-receive-channel-messages-withRSC/nodejs/server/api/botController.js#L24) line and put your debugger for local debug.

## Running the sample

**Adding bot UI:**

![App installation](images/1.Install.png)

**Hey command interaction:**

![Permissions](images/3.Interaction.png)

**1 or 2 command interaction:**

![Permissions](images/4.1_and_2_Command_Interaction.png) 

**Adding App to group chat:**

![Adding To Groupchat](images/5.Install_to_GC.png) 

**Group chat interaction with bot without being @mentioned:**

![Group Chat](images/7.1_and_2_Command_Interaction.png) 

**Interacting with the bot in Teams**

Select a channel and enter a message in the channel for your bot.

The bot receives the message without being @mentioned.

## Deploy the bot to Azure

To learn more about deploying a bot to Azure, see [Deploy your bot to Azure](https://aka.ms/azuredeployment) for a complete list of deployment instructions.

## Further reading

- [Bot Framework Documentation](https://docs.botframework.com)
- [Bot Basics](https://docs.microsoft.com/azure/bot-service/bot-builder-basics?view=azure-bot-service-4.0)
- [Azure Bot Service Introduction](https://docs.microsoft.com/azure/bot-service/bot-service-overview-introduction?view=azure-bot-service-4.0)
- [Azure Bot Service Documentation](https://docs.microsoft.com/azure/bot-service/?view=azure-bot-service-4.0)
- [Receive Channel messages with RSC](https://docs.microsoft.com/microsoftteams/platform/bots/how-to/conversations/channel-messages-with-rsc)


<img src="https://pnptelemetry.azurewebsites.net/microsoft-teams-samples/samples/bot-receive-channel-messages-withRSC-nodejs" />

- https://www.npmjs.com/package/botbuilder#installing
- https://github.com/microsoft/BotBuilder-Samples/tree/main/experimental/generation
- https://github.com/microsoft/BotBuilder-Samples/tree/main/samples/javascript_nodejs/44.prompt-for-user-input
- https://github.com/microsoft/BotBuilder-Samples/tree/main/samples/javascript_nodejs/19.custom-dialogs
- https://github.com/microsoft/BotBuilder-Samples/tree/main/samples/javascript_nodejs/17.multilingual-bot
- https://github.com/microsoft/BotBuilder-Samples/tree/main/samples/javascript_nodejs/05.multi-turn-prompt
- https://github.com/microsoft/BotBuilder-Samples/tree/main/samples/javascript_nodejs/06.using-cards
- https://github.com/microsoft/BotBuilder-Samples/tree/main/samples/javascript_nodejs/07.using-adaptive-cards
- https://github.com/microsoft/BotBuilder-Samples/tree/main/samples/javascript_nodejs/15.handling-attachments
- https://github.com/microsoft/BotBuilder-Samples/tree/main/samples/javascript_nodejs/43.complex-dialog
- https://learn.microsoft.com/en-us/azure/bot-service/bot-service-overview?view=azure-bot-service-4.0
- https://learn.microsoft.com/en-us/javascript/api/botbuilder/?view=botbuilder-ts-latest
- https://www.npmjs.com/package/botbuilder-dialogs#learn-more
- https://learn.microsoft.com/en-us/azure/bot-service/bot-service-overview?view=azure-bot-service-4.0
- https://learn.microsoft.com/en-us/javascript/api/botbuilder-dialogs/?view=botbuilder-ts-latest
- https://github.com/Microsoft/botframework-sdk?tab=readme-ov-file
- https://github.com/Microsoft/botbuilder-js
- https://github.com/howdyai/botkit#readme
- https://github.com/howdyai/botkit/tree/main/packages/botbuilder-adapter-slack#readme
- https://github.com/BotBuilderCommunity/botbuilder-community-js/blob/master/libraries/botbuilder-adapter-console/README.md
- https://github.com/BotBuilderCommunity/botbuilder-community-js/blob/master/libraries/botbuilder-adapter-alexa/README.md
- https://github.com/BotBuilderCommunity/botbuilder-community-js/tree/master/samples/adapter-alexa
- https://github.com/BotBuilderCommunity/botbuilder-community-js/blob/master/libraries/botbuilder-dialog-prompts/README.md
- https://github.com/microsoft/Recognizers-Text
- https://github.com/BotBuilderCommunity/botbuilder-community-js/tree/master/samples/dialog-prompts
- https://github.com/BotBuilderCommunity/botbuilder-community-js/blob/master/libraries/botbuilder-storage-mongodb/README.md
- https://github.com/howdyai/botkit/blob/main/packages/docs/index.md
- https://learn.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-authentication?view=azure-bot-service-4.0&tabs=multitenant

---

## Changelog

### Latest Changes (Daily Recap, Notification Queue, Adaptive Card Templating)

#### New Features

- **Daily Recap Card API** (`POST /api/dailyrecap`)
  - Proactively post Adaptive Cards built from Adaptive Cards Templating template + data pairs
  - Server-side template binding via `adaptivecards-templating` SDK (Teams cannot expand templating syntax client-side)
  - Auto-delete sent cards after 5 minutes to avoid channel clutter
  - Caches template/data in Redis for later `Action.Submit` toggling

- **`daily recap` Bot Command**
  - Admin command fetches daily recap card from MyAdmin endpoint and posts to current channel
  - Compact mode trims card to stay under Teams' 25 KB Adaptive Card limit
  - Auto-delete after 5 minutes

- **Notification Queue Consumer** (`server/queue/notificationConsumer.js`)
  - Polls Redis (`notif:queue`) on a configurable interval
  - Per-room edit/coalesce/new send decisions
  - Edit window: within 30 minutes,同一个 `dedup_key` 的多条消息会合并到同一条活动
  - Coalescing: multiple non-deduped messages combined into one activity
  - Fallback to Power Automate webhook if Bot Framework delivery fails
  - Dead-letter queue for failed/expired envelopes
  - Filtering: noisy GitHub events (check_run, workflow_job, star, fork, etc.) redirected to `int-dev-announce` instead of flooding primary channels

- **`!notif` Admin Commands**
  - `!notif status` — queue/processing/dead depths and metrics
  - `!notif rooms` — list known rooms with ✅/❌ convref status
  - `!notif test <room> <msg>` — enqueue a probe envelope
  - `!notif drain-dead` — re-queue everything in `notif:dead`
  - `!notif seed-room <room>` — check/seed conversation reference for a room

- **Channel Sync on Startup**
  - New `CHANNEL_SYNC_ENABLED=1` env var enables proactive probing of all channels on startup
  - Bot sends a sync-check message to each channel that lacks a stored `convref`
  - Ensures channels added after bot deploy can receive proactive messages

- **Constructed ConversationReference Fallback**
  - When no `convref` is stored for a room, notification consumer attempts to send using a constructed reference
  - Uses standard Teams service URL (`https://smba.trafficmanager.net/teams/`)
  - Falls back to webhook if even the constructed reference fails

#### Refactoring

- **Shared Retry Library** (`server/lib/retry.js`)
  - Centralized retry logic replacing inline loops in `botController.js` and `msgController.js`
  - Exponential backoff with jitter (base 750ms, cap 8s)
  - Catches transient errors, auth failures, and rate limits
  - Re-trusts `serviceUrl` on auth failures (original behavior preserved)

- **Channel-Based Routing**
  - `msgController.js` now accepts `channel` in request body and resolves via `CHANNELS` map
  - Hardcoded `int-dev-private` removed — channel is now configurable per-request
  - `SKIP_CHANNELS` array allows disabling channels without code changes

- **BotFrameworkAdapter Instances**
  - `msgController.js`, `dailyRecapController.js`, and `botActivityHandler.js` each create their own adapter instances
  - `botActivityHandler` exports its adapter for use by startup sync

#### Infrastructure

- **Dual Redis Configuration**
  - Notification queue consumer uses separate Redis connection (`REDIS_HOST_MY`/`REDIS_PORT_MY`)
  - Bot's `convref:*` keys still live on primary Redis (`REDIS_HOST`/`REDIS_PORT`)
  - Supports authenticated Redis for bot Redis, unauthenticated for queue Redis

- **New Environment Variables**
  - `DAILY_RECAP_URL` — MyAdmin daily recap card endpoint
  - `DAILY_RECAP_TOKEN` — Shared secret for daily recap auth
  - `NOTIF_POLL_MS` — Queue poll interval (default 5000ms)
  - `NOTIF_POLL_FAST_MS` — Fast poll interval after activity (default 1000ms)
  - `NOTIF_MAX_PER_TICK` — Max envelopes per poll (default 50)
  - `NOTIF_COALESCE_MAX_CHARS` — Max chars per coalesced message (default 24000)
  - `NOTIF_COALESCE_MAX_ITEMS` — Max items per coalesced message (default 8)
  - `NOTIF_EDIT_WINDOW_MS` — Dedup edit window (default 30 minutes)
  - `NOTIF_KEY_PREFIX` — Redis key prefix for queue (default `notif:`)
  - `NOTIF_HEARTBEAT_MS` — Heartbeat log interval (default 60000ms)
  - `NOTIF_FILTER_ENABLED` — Set to `0` to bypass GitHub noise filters
  - `NOTIF_CONSUMER_ENABLED` — Set to `0` to disable notification consumer
  - `CHANNEL_SYNC_ENABLED` — Set to `1` to enable channel sync on startup
  - `REDIS_HOST_MY` / `REDIS_PORT_MY` — Notification queue Redis host/port
  - `REDIS_USER` / `REDIS_PASSWORD` — Bot Redis authentication (optional)

- **`/health/queue` Endpoint**
  - Returns queue depth, processing depth, dead count, poll interval, and last tick stats
  - Useful for monitoring and alerting
