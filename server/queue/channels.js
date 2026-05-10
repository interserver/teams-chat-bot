// Friendly room name → Microsoft Teams conversation ID.
//
// Mirrors the keys used by MyAdmin's getChatRooms() in include/logging.php
// and the channel set in webhooks.interserver.net/src/config.php so a
// single name routes consistently across all three code bases.
//
// Aliases collapse multiple producer names onto one Teams conversation
// (e.g. MyAdmin's "hardware" and the bot's "int-hw" point at the same
// channel). Add new entries here rather than scattering UUIDs through
// individual controllers.

const CHANNELS = {
    'notifications':     '19:028421460efc48f89e00d1c7217bad63@thread.v2',
    'bot-testing':       '19:f72d44d5c53745c89ed1bfd8cd957fd8@thread.v2',
    'int-dev':           '19:0c93975aae904b7db892891da3065c33@thread.v2',
    'int-dev-private':   '19:0c93975aae904b7db892891da3065c33@thread.v2',
    'int-development':   '19:8VXsuLOoLvQgxWlaCOPEXZUE5vvx-tDWMjQErha-4LI1@thread.v2',
    'int-hw':            '19:LqpEjZTwVYBrsDVZJvrqIBZV-GUSg4rqj2nBWPksfCU1@thread.v2',
    'hardware':          '19:LqpEjZTwVYBrsDVZJvrqIBZV-GUSg4rqj2nBWPksfCU1@thread.v2',
    'general':           '19:028421460efc48f89e00d1c7217bad63@thread.v2',
    'development':       '19:8VXsuLOoLvQgxWlaCOPEXZUE5vvx-tDWMjQErha-4LI1@thread.v2',
    'int-dev-announce':  '19:2e7eb459fb9d4f4eafba85f7f373e71b@thread.v2',
    'interserver.net':   '19:d9dc2f7195f84637b748bc36622612fc@thread.v2'
};

function resolve(roomName) {
    if (!roomName) return null;
    return CHANNELS[roomName] || null;
}

function knownRooms() {
    return Object.keys(CHANNELS);
}

module.exports = { CHANNELS, resolve, knownRooms };
