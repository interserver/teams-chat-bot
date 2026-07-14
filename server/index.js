// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const path = require('path');
const express = require('express');
const dotenv = require('dotenv');
const { rateLimit } = require('express-rate-limit');

// Load environment variables from .env file
const ENV_FILE = path.join(__dirname, '../.env');
dotenv.config({ path: ENV_FILE, quiet: true });

// Validate required env vars before anything else
const { validateEnv } = require('./validateEnv');
validateEnv();

const server = express();

// ---------------------------------------------------------------------------
// Rate limiting — protect API endpoints from abuse.
// Configurable via env vars; defaults are conservative.
// ---------------------------------------------------------------------------
const apiLimiter = rateLimit({
    windowMs: parseInt(process.env.RATE_LIMIT_WINDOW_MS || '60000', 10),  // 1 minute
    limit: parseInt(process.env.RATE_LIMIT_MAX || '60', 10),               // 60 req/min per IP
    standardHeaders: true,
    legacyHeaders: false,
    validate: { xForwardedForHeader: false },  // Don't error when X-Forwarded-For is set but proxy isn't trusted
    message: { error: 'Too many requests, please try again later.' }
});

// Note: the stricter per-channel send limiter lives in server/api/index.js,
// applied directly to the proactive-send routes.

const { startConsumer, stopConsumer, getHealth: getQueueHealth } = require('./queue/notificationConsumer');
const { botActivityHandler } = require('./api/botController');

// Apply general rate limiting to /api
server.use('/api', apiLimiter, require('./api'));

// Health check endpoint
server.get('/health', (_req, res) => {
    res.json({ status: 'ok', uptime: process.uptime() });
});

// Notification queue consumer health
server.get('/health/queue', async (_req, res) => {
    try {
        const data = await getQueueHealth();
        res.json({ status: 'ok', ...data });
    } catch (err) {
        res.status(500).json({ status: 'error', error: err.message });
    }
});

// Handle undefined routes (Express 5: bare '*' is no longer a valid path —
// use a final catch-all middleware instead)
server.use((req, res) => {
    res.status(404).json({ error: 'Route not found' });
});

// Set the port from environment variables or default to 3978
const port = process.env.PORT || 3978;

// Start the server
const httpServer = server.listen(port, () => {
    console.log(`Bot/ME service listening at http://localhost:${ port }`);
    startConsumer();
    // Sync conversation references for channels that were added after bot deploy
    // Disabled by default; set CHANNEL_SYNC_ENABLED=1 to re-enable
    if (process.env.CHANNEL_SYNC_ENABLED === '1') {
        botActivityHandler.syncConversationReferences();
    } else {
        console.log('[startup] channel sync disabled (set CHANNEL_SYNC_ENABLED=1 to enable)');
    }
});

// Graceful shutdown
async function shutdown(signal) {
    console.log(`${ signal } received — shutting down gracefully`);
    try {
        await stopConsumer();
    } catch (err) {
        console.warn('stopConsumer failed:', err.message);
    }
    httpServer.close(() => {
        console.log('HTTP server closed');
        process.exit(0);
    });
    // Force exit after 10 seconds if connections don't close
    setTimeout(() => {
        console.error('Forcefully shutting down');
        process.exit(1);
    }, 10000).unref();
}

process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('SIGINT', () => shutdown('SIGINT'));
