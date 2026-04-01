const { describe, it, beforeEach, afterEach } = require('node:test');
const assert = require('node:assert/strict');
const { validateEnv, REQUIRED } = require('../server/validateEnv');

describe('validateEnv', () => {
    let saved;

    beforeEach(() => {
        saved = {};
        for (const key of REQUIRED) {
            saved[key] = process.env[key];
            process.env[key] = 'test-value';
        }
    });

    afterEach(() => {
        for (const key of REQUIRED) {
            if (saved[key] === undefined) {
                delete process.env[key];
            } else {
                process.env[key] = saved[key];
            }
        }
    });

    it('does not throw when all required vars are set', () => {
        assert.doesNotThrow(() => validateEnv());
    });

    it('throws when a required var is missing', () => {
        delete process.env.MicrosoftAppId;
        assert.throws(() => validateEnv(), /MicrosoftAppId/);
    });

    it('lists all missing vars in the error message', () => {
        delete process.env.MicrosoftAppId;
        delete process.env.MYSQL_HOST;
        assert.throws(() => validateEnv(), /MicrosoftAppId.*MYSQL_HOST|MYSQL_HOST.*MicrosoftAppId/);
    });
});
