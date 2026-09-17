// Shared Redis client factory.
//
// Centralises host/port/auth so that all modules use the same
// environment variables with consistent defaults and helpers.

const Redis = require('ioredis');

// -----------------------------------------------------------------------
// Bot Redis (convref: keys) — used by botActivityHandler, msgController,
// dailyRecapController.  May require auth.
// -----------------------------------------------------------------------
let _botRedis = null;

function buildBotRedis(opts = {}) {
    const host = process.env.REDIS_HOST || '67.217.60.234';
    const port = parseInt(process.env.REDIS_PORT || '6379', 10);
    const cfg = {
        host,
        port,
        lazyConnect: false,
        maxRetriesPerRequest: 3,
        retryStrategy: (tries) => Math.min(tries * 200, 2000),
        ...opts
    };
    if (process.env.REDIS_USER)  cfg.username = process.env.REDIS_USER;
    if (process.env.REDIS_PASSWORD || process.env.REDIS_PASS) {
        cfg.password = process.env.REDIS_PASSWORD || process.env.REDIS_PASS;
    }
    return cfg;
}

function getBotRedis() {
    if (!_botRedis) {
        const cfg = buildBotRedis();
        _botRedis = new Redis(cfg);
        _botRedis.on('connect', () => console.log('✅ Connected to bot Redis'));
        _botRedis.on('error', (err) => console.error('❌ Bot Redis error:', err.message));
    }
    return _botRedis;
}

function createBotRedis(opts = {}) {
    return new Redis(buildBotRedis(opts));
}

// -----------------------------------------------------------------------
// Notification-queue Redis — separate host, no auth by default.
// -----------------------------------------------------------------------
let _notifRedis = null;

function buildNotifRedisOpts(opts = {}) {
    const host = process.env.REDIS_HOST_MY || '67.217.60.234';
    const port = parseInt(process.env.REDIS_PORT_MY || '6379', 10);
    return {
        host,
        port,
        lazyConnect: false,
        maxRetriesPerRequest: 3,
        retryStrategy: (tries) => Math.min(tries * 200, 2000),
        ...opts
    };
}

function getNotifRedis() {
    if (!_notifRedis) {
        _notifRedis = new Redis(buildNotifRedisOpts());
        _notifRedis.on('connect', () => console.log('[notif] ✅ Connected to notif Redis'));
        _notifRedis.on('error', (err) => console.error('[notif] ❌ Notif Redis error:', err.message));
    }
    return _notifRedis;
}

function createNotifRedis(opts = {}) {
    return new Redis(buildNotifRedisOpts(opts));
}

// -----------------------------------------------------------------------
// Teams service URL — GCC/GCC High differ; make it configurable.
// -----------------------------------------------------------------------
const TEAMS_SERVICE_URL = process.env.TEAMS_SERVICE_URL
    || 'https://smba.trafficmanager.net/teams/';

module.exports = {
    getBotRedis,
    createBotRedis,
    getNotifRedis,
    createNotifRedis,
    TEAMS_SERVICE_URL,
    // re-export so callers can use Redis directly if needed
    Redis
};
