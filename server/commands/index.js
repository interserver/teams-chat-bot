// Command registry — maps command matchers to handler functions.
// Each command exports: { match(text, lcText, context), execute(matchResult, deps) }
// deps = { context, member, email, ima, db, redis, usersCollection, execFileAsync, bot }

const commands = [
    require('./ima'),
    require('./ping'),
    require('./joke'),
    require('./setMaster'),
    require('./ticketCard'),
    require('./ticketSubmit'),
    require('./ticketQuick'),
    require('./mailbabyUser'),
    require('./ipLookup'),
    require('./blockEmail'),
    require('./blockDomain'),
    require('./blockHelp'),
    require('./githubIssues'),
    require('./githubLabels'),
    require('./assetSearch'),
    require('./hypervStatus'),
    require('./processingStatus'),
    require('./globalVar'),
    require('./dailyRecap'),
    require('./notifAdmin'),
    require('./help')
];

/**
 * Try each registered command in order.
 * Returns true if a command matched and executed, false otherwise.
 */
async function dispatch(text, lcText, deps) {
    for (const cmd of commands) {
        const matchResult = cmd.match(text, lcText, deps);
        if (matchResult) {
            await cmd.execute(matchResult, deps);
            return true;
        }
    }
    return false;
}

module.exports = { dispatch, commands };
