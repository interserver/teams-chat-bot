// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const path = require('path');
const express = require('express');
const dotenv = require('dotenv');

// Load environment variables from .env file
const ENV_FILE = path.join(__dirname, '../.env');
dotenv.config({ path: ENV_FILE });

// Validate required env vars before anything else
const { validateEnv } = require('./validateEnv');
validateEnv();

const server = express();

// Use the API routes
server.use('/api', require('./api'));

// Health check endpoint
server.get('/health', (_req, res) => {
    res.json({ status: 'ok', uptime: process.uptime() });
});

// Handle undefined routes
server.get('*', (req, res) => {
    res.status(404).json({ error: 'Route not found' });
});

// Set the port from environment variables or default to 3978
const port = process.env.PORT || 3978;

// Start the server
const httpServer = server.listen(port, () => {
    console.log(`Bot/ME service listening at http://localhost:${ port }`);
});

// Graceful shutdown
function shutdown(signal) {
    console.log(`${ signal } received — shutting down gracefully`);
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
