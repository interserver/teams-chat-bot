// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const express = require('express');

const router = express.Router();

// Built-in body parsers (since Express 4.16+)
router.use(express.json()); // for application/json
router.use(express.urlencoded({ extended: true })); // for application/x-www-form-urlencoded

// Route to handle incoming messages
router.post('/messages', require('./botController'));
router.post('/message', require('./msgController'));

module.exports = router;
