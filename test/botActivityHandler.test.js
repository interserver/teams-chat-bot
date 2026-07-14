const { describe, it, after } = require('node:test');
const assert = require('node:assert/strict');

// Test the utility methods extracted to the BotActivityHandler class.
// We can't instantiate BotActivityHandler (needs DB connections) so we
// test the static-like validation methods directly by importing and calling them.

// isValidIP and isValidHostname are instance methods on the prototype,
// so we can call them with a dummy `this`.
const { BotActivityHandler } = require('../server/bot/botActivityHandler');
const proto = BotActivityHandler.prototype;

// After tests, forcibly exit since the module creates DB connections at load time
// (via botController.js singleton) that keep the process alive.
after(() => {
    setTimeout(() => process.exit(0), 100);
});

describe('isValidIP', () => {
    it('accepts valid IPv4', () => {
        assert.ok(proto.isValidIP('192.168.1.1'));
        assert.ok(proto.isValidIP('0.0.0.0'));
        assert.ok(proto.isValidIP('255.255.255.255'));
    });
    it('rejects invalid IPv4', () => {
        assert.equal(proto.isValidIP('999.1.1.1'), false);
        assert.equal(proto.isValidIP('abc'), false);
        assert.equal(proto.isValidIP(''), false);
    });
    it('accepts IPv6 loopback', () => {
        assert.ok(proto.isValidIP('::1'));
    });
    it('accepts full IPv6', () => {
        assert.ok(proto.isValidIP('2001:0db8:85a3:0000:0000:8a2e:0370:7334'));
    });
});

describe('isValidHostname', () => {
    it('accepts valid hostnames', () => {
        assert.ok(proto.isValidHostname('example.com'));
        assert.ok(proto.isValidHostname('sub.domain.example.com'));
        assert.ok(proto.isValidHostname('my-host'));
    });
    it('rejects invalid hostnames', () => {
        assert.equal(proto.isValidHostname('-bad.com'), false);
        assert.equal(proto.isValidHostname(''), false);
        assert.equal(proto.isValidHostname('a'.repeat(254)), false);
    });
});

describe('updateActionSubmitData', () => {
    it('sets activityId on Action.Submit elements', () => {
        const card = {
            type: 'ActionSet',
            actions: [
                { type: 'Action.Submit', data: {} },
                { type: 'Action.OpenUrl', data: {} }
            ]
        };
        proto.updateActionSubmitData(card, { id: 'test-123' });
        assert.equal(card.actions[0].data.activityId, 'test-123');
        assert.equal(card.actions[1].data.activityId, undefined);
    });

    it('recurses into items, columns, and body', () => {
        const card = {
            body: [{
                columns: [{
                    items: [{
                        type: 'ActionSet',
                        actions: [{ type: 'Action.Submit', data: {} }]
                    }]
                }]
            }]
        };
        proto.updateActionSubmitData(card, { id: 'nested-456' });
        assert.equal(card.body[0].columns[0].items[0].actions[0].data.activityId, 'nested-456');
    });
});
