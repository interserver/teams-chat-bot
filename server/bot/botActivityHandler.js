// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TurnContext, TeamsInfo, TeamsActivityHandler, MessageFactory } = require('botbuilder');
const { execFile } = require('child_process');
const { promisify } = require('util');
const fs = require('fs');
const path = require('path');
const Redis = require('ioredis');
const { MongoClient } = require('mongodb');
const mysql = require('mysql2/promise');

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
        const db = mysql.createPool({
            host: process.env.MYSQL_HOST,
            user: process.env.MYSQL_USER,
            password: process.env.MYSQL_PASS,
            database: process.env.MYSQL_DB,
            waitForConnections: true,
            connectionLimit: 10,
            queueLimit: 0
        });
        db.getConnection()
            .then((conn) => {
                console.log("✅ Connected to MySQL");
                conn.release();
            })
            .catch((err) => console.error("❌ MySQL error:", err));

        // Redis
        this.redis = new Redis({ host: 'dragonfly.mailbaby.net', port: 6379 });
        this.redis.on('connect', () => console.log("✅ Connected to Redis"));
        this.redis.on('error', (err) => console.error("❌ Redis error:", err));

        // MongoDB
        const mongoClient = new MongoClient(`mongodb://${ encodeURIComponent(process.env.ZONEMTA_USERNAME) }:${ encodeURIComponent(process.env.ZONEMTA_PASSWORD) }@${ process.env.ZONEMTA_HOST }:27017/`);
        mongoClient.connect()
            .then(() => console.log("✅ Connected to MongoDB"))
            .catch((err) => console.error("❌ MongoDB error:", err));
        const usersCollection = mongoClient.db('zone-mta').collection('users');

        const execFileAsync = promisify(execFile);

        // Predefine regex
        this.hostRegex = /(?<host>([a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,})/;
        this.emailRegex = /(?<email>[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/;

        // Activity called when there's a message in channel
        this.onMessage(async (context, next) => {

            // conversationReference includes serviceUrl, user, bot, and conversation.id
            const conversationReference = TurnContext.getConversationReference(context.activity);
            //console.log(conversationReference);

            // store it somewhere (Redis, Mongo, MySQL, etc.)
            await this.redis.set(`convref:${conversationReference.conversation.id}`, JSON.stringify(conversationReference));

            /* context.activity:
            {
              text: '1',
              textFormat: 'plain',
              attachments: [ { contentType: 'text/html', content: '<p>1</p>' } ],
              type: 'message',
              id: '1755495760922',
              channelId: 'msteams',
              serviceUrl: 'https://smba.trafficmanager.net/amer/7837f8c1-952c-4ce2-8d54-6ef28f5cdfef/',
              from: {
                id: '29:1Tg_ijIGbNJjZZpPViCdFXG7adkoazSTdXmr79tYDeQtivdqiCe4Q3DDbUdxCriDHjzDMfTG-xBWAY971Kr9aRA',
                name: 'Joe Huss',
                aadObjectId: 'c64d7672-0cd2-4d0d-9a9f-7bac202f49b4'
              },
              conversation: {
                conversationType: 'personal',   channel / personal / groupChat
                tenantId: '7837f8c1-952c-4ce2-8d54-6ef28f5cdfef',
                id: 'a:1sU2r-PV8_UqBYLCy4cplte8M37bxk6frzVlrNJjc7CJmLzqWtT07Ll9bgdMaj_47UPUd7vBolBlmA3vTnfiz6-KLZHy-TTeyG2-oQ3MbAXibkX-4KJpZGRInAU2V4HpC'
              },
              recipient: {
                id: '28:6fa7ed27-9923-4d5a-9f2d-2c9b81cfdd2d',
                name: 'InterestingGuy'
              },
              channelData: { tenant: { id: '7837f8c1-952c-4ce2-8d54-6ef28f5cdfef' } },
              locale: 'en-US',
              localTimezone: 'America/New_York',
              callerId: 'urn:botframework:azure'
            }
            */
            const text = (context.activity.text || '').trim();
            const lcText = text.toLowerCase();
            const userId = context.activity.from.id;
            /* member:
            {
              id: '29:1Tg_ijIGbNJjZZpPViCdFXG7adkoazSTdXmr79tYDeQtivdqiCe4Q3DDbUdxCriDHjzDMfTG-xBWAY971Kr9aRA',
              name: 'Joe Huss',
              aadObjectId: 'c64d7672-0cd2-4d0d-9a9f-7bac202f49b4',
              givenName: 'Joe',
              surname: 'Huss',
              email: 'detain@interserver.net',
              userPrincipalName: 'detain@interserver.net',
              tenantId: '7837f8c1-952c-4ce2-8d54-6ef28f5cdfef',
              userRole: 'user'
            }
            */
            //const teamDetails = await TeamsInfo.getTeamDetails(context);
/*            if (teamDetails) {
                // Sends a message activity to the sender of the incoming activity.
                await context.sendActivity(MessageFactory.text(`The group ID is: ${teamDetails.aadGroupId}`));
            } else {
                await context.sendActivity(MessageFactory.text('This message did not come from a channel in a team.'));
            } */
            //const teamChannels = await TeamsInfo.getTeamChannels(context);
            //const members = await TeamsInfo.getMembers(context);
            //console.log(members);
            //console.log(teamDetails);
            //console.log(teamChannels);
            const member = await TeamsInfo.getMember(context, userId);
            const email = member.email || member.userPrincipalName;
            //const teamId = await TeamsInfo.getTeamId(context);
            //console.log(`team id ${teamId}`);
            // console.log('User email:', email);
            // console.log('Conversation:', context.activity.conversation);
            //console.log('Activity:');
            const channelId = context.activity.channelData.tenant.id;
            //console.log(context);
            //console.log(context.activity);
            // console.log('Member:');
            //console.log(member);
            const [accountRow] = await db.query('select * from accounts where account_lid=?', [email]);
            const ima = !accountRow || accountRow.length === 0 ? 'unknown' : accountRow[0].account_ima;
            console.log(`#${channelId} [${ima}] ${member.name} <${email}> sent message: ${text}`);
            let match;
            if (text === 'ima') {
                await context.sendActivity(MessageFactory.text(`Hello ${ member.name }, I see your email is ${ email } and you are ima ${ ima }`));
            } else if ((match = text.match(/^ping\s+(.+)$/i))) {
                const target = match[1].trim();

                if (this.isValidHostname(target) || this.isValidIP(target)) {
                    await context.sendActivity(MessageFactory.text(`Pinging \`${ target }\` ...`));
                    try {
                        // Run ping safely and await result
                        const args = ['-w', '10', '-W', '10', '-c', '4', '-q', target];
                        const { stdout } = await execFileAsync('ping', args, { timeout: 15000 });

                        // Only keep last 3 lines
                        const lines = stdout.trim().split('\n');
                        const output = lines.slice(-3).join('\n');
                        await context.sendActivity(MessageFactory.text('```\n' + output + '\n```'));
                    } catch (err) {
                        await context.sendActivity(MessageFactory.text(`⚠️ Error: ${ err.stderr || err.message }`));
                    }
                } else {
                    await context.sendActivity(MessageFactory.text(`❌ Invalid hostname or IP: \`${ target }\``));
                }
            } else if (lcText === 'joke' || lcText === 'tell a joke') {
                try {
                    const jokesPath = path.join(__dirname, '../../jokes.json');
                    const jokes = JSON.parse(fs.readFileSync(jokesPath, 'utf8'));
                    const jokeList = Object.values(jokes).flat();
                    if (Array.isArray(jokeList) && jokeList.length > 0) {
                        const randomJoke = jokeList[Math.floor(Math.random() * jokeList.length)];
                        if (Array.isArray(randomJoke)) {
                            // Send each line separately
                            for (const line of randomJoke) {
                                await context.sendActivity(MessageFactory.text(line));
                            }
                        } else {
                            // Fallback if joke isn’t an array
                            await context.sendActivity(MessageFactory.text(String(randomJoke)));
                        }
                    } else {
                        await context.sendActivity(MessageFactory.text('Hmm, I don’t have any jokes right now 😅'));
                    }
                } catch (err) {
                    console.error('Error loading jokes.json:', err);
                    await context.sendActivity(MessageFactory.text('⚠️ Sorry, I couldn’t fetch a joke right now.'));
                }
            } else if (ima === 'admin') {
                if ((match = lcText.match(/^add mailbaby user (\S+) (\S+)$/i))) {
                    const user = match[1];
                    const pass = match[2];
                    const existing = await usersCollection.findOne({ username: user });
                    if (existing) {
                        await context.sendActivity(MessageFactory.text(`Found existing user '${ user }'`));
                    } else {
                        const result = await usersCollection.insertOne({ username: user, password: pass });
                        if (result.insertedId) {
                            await context.sendActivity(MessageFactory.text(`Added user '${ user }' with password '${ pass }'`));
                        } else {
                            await context.sendActivity(MessageFactory.text(`Error adding user '${ user }' with password '${ pass }'`));
                        }
                    }
                } else if ((match = text.match(/^delete mailbaby user (\S+)$/i))) {
                    const user = match[1];
                    const existing = await usersCollection.findOne({ username: user });
                    if (existing) {
                        await usersCollection.deleteOne({ username: user });
                        await context.sendActivity(MessageFactory.text(`Removed user '${ user }'`));
                    } else {
                        await context.sendActivity(MessageFactory.text(`No user '${ user }' exists`));
                    }
                } else if ((match = text.match(/.*(where|lookup|query|find|locate|search).*?[^\d](\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})[^\d]?.*/i))) {
                    const ip = match[2];
                    const [rows] = await db.query(
                        `SELECT *, assets.id AS real_asset_id
                        FROM ips
                        LEFT JOIN vlans ON ips_vlan=vlans_id
                        LEFT JOIN switchports ON FIND_IN_SET(ips_vlan, vlans) != 0
                        LEFT JOIN assets ON switchports.asset_id=assets.id
                        LEFT JOIN asset_types ON type_id=asset_types.asset_id
                        LEFT JOIN asset_locations ON location_id=datacenter
                        LEFT JOIN asset_racks ON rack_id=rack
                        LEFT JOIN switchmanager ON switchports.switch=switchmanager.id
                        WHERE ips_ip = ?`,
                        [ip]
                    );

                    if (!rows || rows.length === 0) {
                        await context.sendActivity(MessageFactory.text(`Unable to find ${ ip } in our IP database`));
                    } else {
                        const r = rows[0];
                        const unit = r.unit_start !== r.unit_end ? `${ r.unit_start }-${ r.unit_end }` : r.unit_start;

                        await context.sendActivity(`Asset: ${ r.real_asset_id }\n
    Hostname: ${ r.hostname }\n
    Status: ${ r.status }\n
    Rack: ${ r.location_name } ${ r.rack_name }\n
    Unit: ${ unit }\n
    Network: Switch${ r.name } Port ${ r.port }\n
    https://my.interserver.net/admin/view_server_order?id=${ r.order_id }\n
    https://my.interserver.net/admin/asset_form?id=${ r.real_asset_id }`
                        );
                    }
                } else if (/^(blocks|block list|blocks list)$/i.test(lcText)) {
                    const blockedEmails = await this.redis.smembers('blocked_emails');
                    blockedEmails.sort();
                    const text = `*Blocked Emails* (${ blockedEmails.length })\n` + blockedEmails.join(', ');
                    await context.sendActivity(MessageFactory.text(text));
                } else if (new RegExp(`^(block|block email) ${ this.emailRegex.source }$`, 'i').test(lcText)) {
                    const match = lcText.match(new RegExp(this.emailRegex, 'i'));
                    const email = match.groups.email;
                    const added = await this.redis.sadd('blocked_emails', email);
                    if (added) {
                        await context.sendActivity(MessageFactory.text(`✅ Successfully added *${ email }* to blocked emails list.`));
                    } else {
                        await context.sendActivity(MessageFactory.text(`⚠️ *${ email }* already exists in blocked emails list.`));
                    }
                } else if (new RegExp(`^(block remove|block delete|block email remove|block email delete|blocked email remove|blocked email delete) ${ this.emailRegex.source }$`, 'i').test(lcText)) {
                    const match = lcText.match(new RegExp(this.emailRegex, 'i'));
                    const email = match.groups.email;
                    const removed = await this.redis.srem('blocked_emails', email);
                    if (removed) {
                        await context.sendActivity(MessageFactory.text(`✅ Successfully removed *${ email }* from blocked emails list.`));
                    } else {
                        await context.sendActivity(MessageFactory.text(`⚠️ *${ email }* is not in blocked emails list.`));
                    }
                } else if (/^(blocked domains|blocked hosts|blocked domains list|blocked hosts list)$/i.test(lcText)) {
                    const blockedDomains = await this.redis.smembers('blocked_domains');
                    blockedDomains.sort();
                    const text = `*Blocked Domains* (${ blockedDomains.length })\n` + blockedDomains.join(', ');
                    await context.sendActivity(MessageFactory.text(text));
                } else if (new RegExp(`^(block domain|block hostname|block host) ${ this.hostRegex.source }$`, 'i').test(lcText)) {
                    const match = lcText.match(new RegExp(this.hostRegex, 'i'));
                    const host = match.groups.host;
                    const added = await this.redis.sadd('blocked_domains', host);
                    if (added) {
                        await context.sendActivity(MessageFactory.text(`✅ Successfully added *${ host }* to blocked domains list.`));
                    } else {
                        await context.sendActivity(MessageFactory.text(`⚠️ *${ host }* already exists in blocked domains list.`));
                    }
                } else if (new RegExp(`^(block remove domain|block delete domain|block domain remove|block domain delete|blocked domain remove|blocked domain delete|blocked domains remove|blocked domains delete) ${ this.hostRegex.source }$`, 'i').test(lcText)) {
                    const match = lcText.match(new RegExp(this.hostRegex, 'i'));
                    const host = match.groups.host;
                    const removed = await this.redis.srem('blocked_domains', host);
                    if (removed) {
                        await context.sendActivity(MessageFactory.text(`✅ Successfully removed *${ host }* from blocked domains list.`));
                    } else {
                        await context.sendActivity(MessageFactory.text(`⚠️ *${ host }* is not in blocked domains list.`));
                    }
                } else if (/^blocks? help$/i.test(lcText)) {
                    const commands = {
                        'blocks list': { description: 'List all blocked Emails' },
                        'block <email>': { description: 'Adds an email to the blocked emails list.' },
                        'block remove <email>': { description: 'Removes an email address from the blocked emails list.' },
                        'blocked domains': { description: 'List all blocked Domains' },
                        'block domain <host>': { description: 'Adds a domain to the blocked domains list.' },
                        'block domain remove <host>': { description: 'Removes a domain from the blocked domains list.' },
                        'block help': { description: 'Show all available Blocked Email/Domains commands' }
                    };
                    let text = '*MailBaby Blocked Emails and Domains Help*\n';
                    for (const [command, details] of Object.entries(commands)) {
                        text += `\`${ command }\` - ${ details.description }\n`;
                    }
                    await context.sendActivity(MessageFactory.text(text));
                }
            }
            await next();
        });

        // Called when the bot is added to a team.
        this.onMembersAdded(async (context, next) => {
            var welcomeText = 'Hello and welcome! With this sample your bot can receive user messages across standard channels in a team without being @mentioned';
            await context.sendActivity(MessageFactory.text(welcomeText));
            await next();
        });

        this.onCommand(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onCommand event'));
            console.log('got onCommand event');
            console.log(context.activity);
            await next();
        });

        this.onCommandResult(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onCommandResult event'));
            console.log('got onCommandResult event');
            console.log(context.activity);
            await next();
        });

        this.onConversationUpdate(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onConversationUpdate event'));
            console.log('got onConversationUpdate event');
            console.log(context.activity);
            await next();
        });

        /*this.onDialog(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onDialog event'));
            console.log('got onDialog event');
            console.log(context.activity);
            await next();
        });*/

        this.onEndOfConversation(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onEndOfConversation event'));
            console.log('got onEndOfConversation event');
            console.log(context.activity);
            await next();
        });

        this.onEvent(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onEvent event'));
            console.log('got onEvent event');
            console.log(context.activity);
            await next();
        });

        this.onInstallationUpdateAdd(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onInstallationUpdateAdd event'));
            console.log('got onInstallationUpdateAdd event');
            console.log(context.activity);
            await next();
        });

        this.onInstallationUpdate(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onInstallationUpdate event'));
            console.log('got onInstallationUpdate event');
            console.log(context.activity);
            await next();
        });

        this.onInstallationUpdateRemove(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onInstallationUpdateRemove event'));
            console.log('got onInstallationUpdateRemove event');
            console.log(context.activity);
            await next();
        });

        this.onMembersRemoved(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onMembersRemoved event'));
            console.log('got onMembersRemoved event');
            console.log(context.activity);
            await next();
        });

        this.onMessageDelete(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onMessageDelete event'));
            console.log('got onMessageDelete event');
            console.log(context.activity);
            await next();
        });

        this.onMessageReaction(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onMessageReaction event'));
            console.log('got onMessageReaction event');
            console.log(context.activity);
            await next();
        });

        this.onMessageUpdate(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onMessageUpdate event'));
            console.log('got onMessageUpdate event');
            console.log(context.activity);
            await next();
        });

        this.onReactionsAdded(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onReactionsAdded event'));
            console.log('got onReactionsAdded event');
            console.log(context.activity);
            await next();
        });

        this.onReactionsRemoved(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onReactionsRemoved event'));
            console.log('got onReactionsRemoved event');
            console.log(context.activity);
            await next();
        });

        this.onTokenResponseEvent(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onTokenResponseEvent event'));
            console.log('got onTokenResponseEvent event');
            console.log(context.activity);
            await next();
        });

        /*
        this.onTurn(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onTurn event'));
            console.log('got onTurn event');
            //console.log(context.activity);
            console.log(context.activity.channelData);
            await next();
        });
        */

        this.onTyping(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onTyping event'));
            console.log('got onTyping event');
            console.log(context.activity);
            await next();
        });

        this.onUnrecognizedActivityType(async (context, next) => {
            //await context.sendActivity(MessageFactory.text('Got onUnrecognizedActivityType event'));
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
}

module.exports.BotActivityHandler = BotActivityHandler;
