// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const express = require('express');
const { rateLimit, ipKeyGenerator } = require('express-rate-limit');

const router = express.Router();

// Built-in body parsers (since Express 4.16+)
router.use(express.json()); // for application/json
router.use(express.urlencoded({ extended: true })); // for application/x-www-form-urlencoded

// Stricter rate limit for proactive-send endpoints
const sendLimiter = rateLimit({
    windowMs: parseInt(process.env.RATE_LIMIT_SEND_WINDOW_MS || '60000', 10),
    limit: parseInt(process.env.RATE_LIMIT_SEND_MAX || '30', 10),
    standardHeaders: true,
    legacyHeaders: false,
    validate: { xForwardedForHeader: false },
    message: { error: 'Send limit exceeded, please slow down.' },
    keyGenerator: (req) => {
        const channel = req.body && req.body.channel ? `:${ req.body.channel }` : '';
        return `${ ipKeyGenerator(req.ip) }${ channel }`;
    }
});

// Route to handle incoming messages
router.post('/messages', require('./botController'));
router.post('/message', sendLimiter, require('./msgController'));
router.post('/dailyrecap', sendLimiter, require('./dailyRecapController'));

module.exports = router;
