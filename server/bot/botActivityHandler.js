// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TurnContext, TeamsInfo, TeamsActivityHandler, MessageFactory } = require('botbuilder');
const { BotFrameworkAdapter } = require('botbuilder');
const { MicrosoftAppCredentials } = require('botframework-connector');
const { execFile } = require('child_process');
const { promisify } = require('util');
const Redis = require('ioredis');
const { MongoClient } = require('mongodb');
const mysql = require('mysql2/promise');
const { dispatch } = require('../commands');
const { CHANNELS } = require('../queue/channels');
const { runWithRetry } = require('../lib/retry');

/*
notifications 19:028421460efc48f89e00d1c7217bad63@thread.v2
bot testing group 19:f72d44d5c53745c89ed1bfd8cd957fd8@thread.v2
int-dev-private 19:0c93975aae904b7db892891da3065c33@thread.v2
int-development 19:8VXsuLOoLvQgxWlaCOPEXZUE5vvx-tDWMjQErha-4LI1@thread.v2
int-hw 19:LqpEjZTwVYBrsDVZJvrqIBZV-GUSg4rqj2nBWPksfCU1@thread.v2
interserver.net 19:d9dc2f7195f84637b748bc36622612fc@thread.v2
*/

class BotActivityHandler extends TeamsActivityHandler {
    constructor() {
        super();

        // MySQL
        this.db = mysql.createPool({
            host: process.env.MYSQL_HOST,
            user: process.env.MYSQL_USER,
            password: process.env.MYSQL_PASS,
            database: process.env.MYSQL_DB,
            waitForConnections: true,
            connectionLimit: 10,
            queueLimit: 0
        });
        this.db.getConnection()
            .then((conn) => {
                console.log('✅ Connected to MySQL');
                conn.release();
            })
            .catch((err) => console.error('❌ MySQL error:', err));

        // Redis
        this.redis = new Redis({ host: process.env.REDIS_HOST || '67.217.60.234', port: parseInt(process.env.REDIS_PORT || '6379', 10) });
        this.redis.on('connect', () => console.log('✅ Connected to Redis'));
        this.redis.on('error', (err) => console.error('❌ Redis error:', err));

        // BotFrameworkAdapter for proactive messages
        this.adapter = new BotFrameworkAdapter({
            appId: process.env.MicrosoftAppId,
            appPassword: process.env.MicrosoftAppPassword
        });
        this.adapter.onTurnError = async (context, error) => {
            console.error(`[sync adapter onTurnError] ${ error }`);
        };

        // MongoDB
        const mongoClient = new MongoClient(`mongodb://${ encodeURIComponent(process.env.ZONEMTA_USERNAME) }:${ encodeURIComponent(process.env.ZONEMTA_PASSWORD) }@${ process.env.ZONEMTA_HOST }:27017/`);
        mongoClient.connect()
            .then(() => console.log('✅ Connected to MongoDB'))
            .catch((err) => console.error('❌ MongoDB error:', err));
        this.usersCollection = mongoClient.db('zone-mta').collection('users');

        this.execFileAsync = promisify(execFile);

        // Predefine regex
        this.hostRegex = /(?<host>([a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,})/;
        this.emailRegex = /(?<email>[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/;

        this.masterTables = { backup_masters: 'backup', website_masters: 'website', vps_masters: 'vps', qs_masters: 'qs' };
        this.tableFields = { backup_masters: [], website_masters: [], vps_masters: [], qs_masters: [] };
        this.detailTables = ['vps_masters', 'qs_masters'];
        for (const table in this.masterTables) {
            const prefix = this.masterTables[table];
            this.db.query('DESCRIBE ??', [table])
                .then(([rows]) => {
                    if (rows && rows.length > 0) {
                        rows.forEach(row => {
                            const field = row.Field.replace(prefix + '_', '');
                            this.tableFields[table].push(field);
                        });
                    }
                }).catch(err => {
                    console.error(`Error describing table ${ table }:`, err);
                });
        }

        // Activity called when there's a message in channel
        this.onMessage(async (context, next) => {
            const conversationReference = TurnContext.getConversationReference(context.activity);
            const convId = conversationReference.conversation.id;
            // Store convref - also covers RSC-enabled channels that missed
            // onInstallationUpdateAdd (e.g., channels added to CHANNELS after bot deploy)
            await this.redis.set(`convref:${ convId }`, JSON.stringify(conversationReference));
            console.log(`[convref] stored for conversation ${ convId } via onMessage`);
            const text = (context.activity.text || '').trim();
            const lcText = text.toLowerCase();
            const userId = context.activity.from.id;

            const member = await TeamsInfo.getMember(context, userId);
            const email = member.email || member.userPrincipalName;
            const channelId = context.activity.channelData.tenant.id;
            const [accountRow] = await this.db.query('select * from accounts where account_lid=?', [email]);
            const ima = !accountRow || accountRow.length === 0 ? 'unknown' : accountRow[0].account_ima;
            console.log(`#${ channelId } [${ ima }] ${ member.name } <${ email }> sent message: ${ text }`);

            // Dispatch to command registry
            const deps = {
                context,
                member,
                email,
                ima,
                db: this.db,
                redis: this.redis,
                usersCollection: this.usersCollection,
                execFileAsync: this.execFileAsync,
                bot: this
            };
            await dispatch(text, lcText, deps);
            await next();
        });

        // Called when the bot is added to a team.
        this.onMembersAdded(async (context, next) => {
            var welcomeText = 'Hello and welcome! With this sample your bot can receive user messages across standard channels in a team without being @mentioned';
            await context.sendActivity(MessageFactory.text(welcomeText));
            await next();
        });

        this.onCommand(async (context, next) => {
            console.log('got onCommand event');
            console.log(context.activity);
            await next();
        });

        this.onCommandResult(async (context, next) => {
            console.log('got onCommandResult event');
            console.log(context.activity);
            await next();
        });

        this.onConversationUpdate(async (context, next) => {
            console.log('got onConversationUpdate event');
            console.log(context.activity);
            await next();
        });

        this.onEndOfConversation(async (context, next) => {
            console.log('got onEndOfConversation event');
            console.log(context.activity);
            await next();
        });

        this.onEvent(async (context, next) => {
            console.log('got onEvent event');
            console.log(context.activity);
            await next();
        });

        this.onInstallationUpdateAdd(async (context, next) => {
            console.log('got onInstallationUpdateAdd event');
            console.log(context.activity);
            // Store ConversationReference so the notification consumer can send
            // proactive messages to this channel without requiring an inbound message first.
            const conversationReference = TurnContext.getConversationReference(context.activity);
            await this.redis.set(`convref:${ conversationReference.conversation.id }`, JSON.stringify(conversationReference));
            console.log(`[convref] stored for conversation ${ conversationReference.conversation.id }`);
            await next();
        });

        this.onInstallationUpdate(async (context, next) => {
            console.log('got onInstallationUpdate event');
            console.log(context.activity);
            await next();
        });

        this.onInstallationUpdateRemove(async (context, next) => {
            console.log('got onInstallationUpdateRemove event');
            console.log(context.activity);
            await next();
        });

        this.onMembersRemoved(async (context, next) => {
            console.log('got onMembersRemoved event');
            console.log(context.activity);
            await next();
        });

        this.onMessageDelete(async (context, next) => {
            console.log('got onMessageDelete event');
            console.log(context.activity);
            await next();
        });

        this.onMessageReaction(async (context, next) => {
            console.log('got onMessageReaction event');
            console.log(context.activity);
            await next();
        });

        this.onMessageUpdate(async (context, next) => {
            console.log('got onMessageUpdate event');
            console.log(context.activity);
            await next();
        });

        this.onReactionsAdded(async (context, next) => {
            console.log('got onReactionsAdded event');
            console.log(context.activity);
            await next();
        });

        this.onReactionsRemoved(async (context, next) => {
            console.log('got onReactionsRemoved event');
            console.log(context.activity);
            await next();
        });

        this.onTokenResponseEvent(async (context, next) => {
            console.log('got onTokenResponseEvent event');
            console.log(context.activity);
            await next();
        });

        this.onTyping(async (context, next) => {
            console.log('got onTyping event');
            console.log(context.activity);
            await next();
        });

        this.onUnrecognizedActivityType(async (context, next) => {
            console.log('got onUnrecognizedActivityType event');
            console.log(context.activity);
            await next();
        });
    }

    // Validate IPv4 + IPv6
    isValidIP(input) {
        const ipv4 = /^(25[0-5]|2[0-4]\d|[0-1]?\d{1,2})(\.(25[0-5]|2[0-4]\d|[0-1]?\d{1,2})){3}$/;
        const ipv6 = /^(([0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}|::1)$/;
        return ipv4.test(input) || ipv6.test(input);
    }

    // Validate hostname
    isValidHostname(input) {
        const hostname = /^(?=.{1,253}$)(?!-)[A-Za-z0-9-]{1,63}(?<!-)(\.(?!-)[A-Za-z0-9-]{1,63}(?<!-))*$/;
        return hostname.test(input);
    }

    updateActionSubmitData(element, sentActivity) {
        if (element.type === 'ActionSet' && Array.isArray(element.actions)) {
            element.actions.forEach(action => {
                if (action.type === 'Action.Submit') {
                    action.data.activityId = sentActivity.id;
                }
            });
        }
        if (Array.isArray(element.items)) {
            element.items.forEach(item => this.updateActionSubmitData(item, sentActivity));
        }
        if (Array.isArray(element.columns)) {
            element.columns.forEach(col => this.updateActionSubmitData(col, sentActivity));
        }
        if (Array.isArray(element.body)) {
            element.body.forEach(child => this.updateActionSubmitData(child, sentActivity));
        }
    }

    async setMaster(context, server, field, value) {
        const { MessageFactory } = require('botbuilder');
        // Input sanitization: restrict server names and field/value length
        const SAFE_NAME = /^[a-zA-Z0-9._-]{1,128}$/;
        const MAX_VALUE_LEN = 256;
        if (!SAFE_NAME.test(server)) {
            await context.sendActivity(MessageFactory.text(`Invalid server name: \`${ server }\``));
            return;
        }
        if (!SAFE_NAME.test(field)) {
            await context.sendActivity(MessageFactory.text(`Invalid field name: \`${ field }\``));
            return;
        }
        if (String(value).length > MAX_VALUE_LEN) {
            await context.sendActivity(MessageFactory.text('Value is too long'));
            return;
        }

        let found = false;
        for (const table in this.masterTables) {
            const prefix = this.masterTables[table];
            const nameField = prefix + '_name';
            try {
                const [rows] = await this.db.query('SELECT * FROM ?? WHERE ?? = ? LIMIT 1', [table, nameField, server]);
                if (rows && rows.length > 0) {
                    server = rows[0][nameField];
                    if (this.tableFields[table].includes(field)) {
                        if (String(rows[0][prefix + '_available']) === String(value)) {
                            await context.sendActivity(MessageFactory.text(`${ server } in ${ table } is already marked available=${ value }`));
                            console.log(`${ server } in ${ table } is already marked available=${ value }`);
                        } else {
                            await this.db.query('UPDATE ?? SET ?? = ? WHERE ?? = ?', [table, prefix + '_' + field, value, nameField, server]);
                            await context.sendActivity(MessageFactory.text(`Updated ${ field }=${ value } in ${ table } where ${ nameField } = ${ server }`));
                            console.log(`Updated ${ field }=${ value } in ${ table } where ${ nameField } = ${ server }`);
                        }
                    } else {
                        await context.sendActivity(MessageFactory.text(`field ${ field } does not exist for ${ server } in ${ table }`));
                        console.log(`field ${ field } does not exist for ${ server } in ${ table }`);
                    }
                    found = true;
                    break;
                }
            } catch (err) {
                await context.sendActivity(MessageFactory.text(`Error querying table ${ table }:`, err));
                console.error(`Error querying table ${ table }:`, err);
            }
        }
        if (!found) {
            await context.sendActivity(MessageFactory.text(`No matching server "${ server }" found in any master table`));
            console.log(`No matching server "${ server }" found in any master table`);
        }
    }

    async syncConversationReferences() {
        const channelEntries = Object.entries(CHANNELS);
        console.log(`[sync] checking ${ channelEntries.length } channels for convrefs...`);

        let probed = 0;
        let hadConvref = 0;

        for (const [channelName, conversationId] of channelEntries) {
            const redisKey = `convref:${ conversationId }`;
            const stored = await this.redis.get(redisKey);

            if (stored) {
                console.log(`[sync] stored convref for ${ channelName } (${ conversationId })`);
                hadConvref++;
                continue;
            }

            console.log(`[sync] missing convref for ${ channelName } (${ conversationId }) - sending probe`);

            try {
                // Build a minimal conversation reference for continueConversation
                const conversationReference = {
                    channelId: 'msteams',
                    conversation: { id: conversationId },
                    serviceUrl: 'https://smba.trafficmanager.net/teams/'
                };

                await runWithRetry(async () => {
                    await this.adapter.continueConversation(conversationReference, async (proactiveContext) => {
                        await MicrosoftAppCredentials.trustServiceUrl(conversationReference.serviceUrl);
                        await proactiveContext.sendActivity('🔄 Sync check');
                    });
                }, {
                    label: 'sync',
                    serviceUrl: conversationReference.serviceUrl,
                    maxRetries: 3
                });

                probed++;
                console.log(`[sync] probe sent to ${ channelName }`);
            } catch (err) {
                console.error(`[sync] failed to probe ${ channelName }: ${ err.message }`);
            }
        }

        console.log(`[sync] complete - ${ probed } channels probed, ${ hadConvref } had convrefs`);
    }
}

module.exports.BotActivityHandler = BotActivityHandler;
