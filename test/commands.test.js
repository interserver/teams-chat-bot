const { describe, it } = require('node:test');
const assert = require('node:assert/strict');

// Command match tests — these only test pattern matching, no I/O needed.

const ima = require('../server/commands/ima');
const ping = require('../server/commands/ping');
const joke = require('../server/commands/joke');
const setMaster = require('../server/commands/setMaster');
const ticketCard = require('../server/commands/ticketCard');
const ticketSubmit = require('../server/commands/ticketSubmit');
const ticketQuick = require('../server/commands/ticketQuick');
const mailbabyUser = require('../server/commands/mailbabyUser');
const ipLookup = require('../server/commands/ipLookup');
const blockEmail = require('../server/commands/blockEmail');
const blockDomain = require('../server/commands/blockDomain');
const blockHelp = require('../server/commands/blockHelp');

// ─── ima ────────────────────────────────────────────────────────────────────

describe('ima command', () => {
    it('matches "ima"', () => {
        assert.ok(ima.match('ima', 'ima'));
    });
    it('does not match other text', () => {
        assert.equal(ima.match('image', 'image'), null);
    });
});

// ─── ping ───────────────────────────────────────────────────────────────────

describe('ping command', () => {
    it('matches "ping example.com"', () => {
        const m = ping.match('ping example.com', 'ping example.com');
        assert.ok(m);
        assert.equal(m.target, 'example.com');
    });
    it('matches "PING 1.2.3.4"', () => {
        const m = ping.match('PING 1.2.3.4', 'ping 1.2.3.4');
        assert.ok(m);
        assert.equal(m.target, '1.2.3.4');
    });
    it('does not match bare "ping"', () => {
        assert.equal(ping.match('ping', 'ping'), null);
    });
});

// ─── joke ───────────────────────────────────────────────────────────────────

describe('joke command', () => {
    it('matches "joke"', () => {
        assert.ok(joke.match('joke', 'joke'));
    });
    it('matches "tell a joke"', () => {
        assert.ok(joke.match('tell a joke', 'tell a joke'));
    });
    it('does not match "joker"', () => {
        assert.equal(joke.match('joker', 'joker'), null);
    });
});

// ─── setMaster ──────────────────────────────────────────────────────────────

describe('setMaster command', () => {
    const deps = { ima: 'admin' };
    it('matches "mark server1 available"', () => {
        const m = setMaster.match('mark server1 available', 'mark server1 available', deps);
        assert.ok(m);
        assert.equal(m.server, 'server1');
        assert.equal(m.field, 'available');
        assert.equal(m.value, '1');
    });
    it('matches "set server1 disabled"', () => {
        const m = setMaster.match('set server1 disabled', 'set server1 disabled', deps);
        assert.ok(m);
        assert.equal(m.value, '0');
    });
    it('matches "set server1 cpu_load to 80"', () => {
        const m = setMaster.match('set server1 cpu_load to 80', 'set server1 cpu_load to 80', deps);
        assert.ok(m);
        assert.equal(m.field, 'cpu_load');
        assert.equal(m.value, '80');
    });
    it('does not match for non-admin', () => {
        assert.equal(setMaster.match('mark server1 available', 'mark server1 available', { ima: 'user' }), null);
    });
});

// ─── ticketCard ─────────────────────────────────────────────────────────────

describe('ticketCard command', () => {
    it('matches "add ticket" for admin', () => {
        assert.ok(ticketCard.match('add ticket', 'add ticket', { ima: 'admin' }));
    });
    it('does not match for non-admin', () => {
        assert.equal(ticketCard.match('add ticket', 'add ticket', { ima: 'user' }), null);
    });
});

// ─── ticketSubmit ───────────────────────────────────────────────────────────

describe('ticketSubmit command', () => {
    const adminDeps = { ima: 'admin', context: { activity: { value: { msteams: { type: 'addTicketSubmit' } } } } };
    it('matches submit action', () => {
        const m = ticketSubmit.match('', '', adminDeps);
        assert.ok(m);
        assert.equal(m.action, 'submit');
    });
    it('matches cancel action', () => {
        const deps = { ima: 'admin', context: { activity: { value: { msteams: { type: 'addTicketCancel' } } } } };
        const m = ticketSubmit.match('', '', deps);
        assert.ok(m);
        assert.equal(m.action, 'cancel');
    });
    it('does not match without msteams value', () => {
        assert.equal(ticketSubmit.match('hello', 'hello', { ima: 'admin', context: { activity: {} } }), null);
    });
});

// ─── ticketQuick ────────────────────────────────────────────────────────────

describe('ticketQuick command', () => {
    it('matches "add billing ticket My subject"', () => {
        const m = ticketQuick.match('add billing ticket My subject', 'add billing ticket my subject', { ima: 'admin' });
        assert.ok(m);
        assert.equal(m.dept, 'billing');
    });
    it('does not match unknown department', () => {
        assert.equal(ticketQuick.match('add foo ticket bar', 'add foo ticket bar', { ima: 'admin' }), null);
    });
});

// ─── mailbabyUser ───────────────────────────────────────────────────────────

describe('mailbabyUser command', () => {
    it('matches add user', () => {
        const m = mailbabyUser.match('add mailbaby user joe pass123', 'add mailbaby user joe pass123', { ima: 'admin' });
        assert.ok(m);
        assert.equal(m.action, 'add');
        assert.equal(m.user, 'joe');
        assert.equal(m.pass, 'pass123');
    });
    it('matches delete user', () => {
        const m = mailbabyUser.match('delete mailbaby user joe', 'delete mailbaby user joe', { ima: 'admin' });
        assert.ok(m);
        assert.equal(m.action, 'delete');
    });
});

// ─── ipLookup ───────────────────────────────────────────────────────────────

describe('ipLookup command', () => {
    it('matches "where is 10.0.0.1"', () => {
        const m = ipLookup.match('where is 10.0.0.1', 'where is 10.0.0.1', { ima: 'admin' });
        assert.ok(m);
        assert.equal(m.ip, '10.0.0.1');
    });
    it('matches "lookup 192.168.1.1 please"', () => {
        const m = ipLookup.match('lookup 192.168.1.1 please', 'lookup 192.168.1.1 please', { ima: 'admin' });
        assert.ok(m);
        assert.equal(m.ip, '192.168.1.1');
    });
    it('does not match for non-admin', () => {
        assert.equal(ipLookup.match('where is 10.0.0.1', 'where is 10.0.0.1', { ima: 'user' }), null);
    });
});

// ─── blockEmail ─────────────────────────────────────────────────────────────

describe('blockEmail command', () => {
    it('matches "blocks list"', () => {
        const m = blockEmail.match('blocks list', 'blocks list', { ima: 'admin' });
        assert.ok(m);
        assert.equal(m.action, 'list');
    });
    it('matches "block test@example.com"', () => {
        const m = blockEmail.match('block test@example.com', 'block test@example.com', { ima: 'admin' });
        assert.ok(m);
        assert.equal(m.action, 'add');
        assert.equal(m.email, 'test@example.com');
    });
    it('matches "block remove test@example.com"', () => {
        const m = blockEmail.match('block remove test@example.com', 'block remove test@example.com', { ima: 'admin' });
        assert.ok(m);
        assert.equal(m.action, 'remove');
    });
});

// ─── blockDomain ────────────────────────────────────────────────────────────

describe('blockDomain command', () => {
    it('matches "blocked domains"', () => {
        const m = blockDomain.match('blocked domains', 'blocked domains', { ima: 'admin' });
        assert.ok(m);
        assert.equal(m.action, 'list');
    });
    it('matches "block domain example.com"', () => {
        const m = blockDomain.match('block domain example.com', 'block domain example.com', { ima: 'admin' });
        assert.ok(m);
        assert.equal(m.action, 'add');
        assert.equal(m.host, 'example.com');
    });
    it('matches "block domain remove example.com"', () => {
        const m = blockDomain.match('block domain remove example.com', 'block domain remove example.com', { ima: 'admin' });
        assert.ok(m);
        assert.equal(m.action, 'remove');
    });
});

// ─── blockHelp ──────────────────────────────────────────────────────────────

describe('blockHelp command', () => {
    it('matches "block help"', () => {
        assert.ok(blockHelp.match('block help', 'block help', { ima: 'admin' }));
    });
    it('matches "blocks help"', () => {
        assert.ok(blockHelp.match('blocks help', 'blocks help', { ima: 'admin' }));
    });
    it('does not match for non-admin', () => {
        assert.equal(blockHelp.match('block help', 'block help', { ima: 'user' }), null);
    });
});
