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
const githubIssues = require('../server/commands/githubIssues');
const githubLabels = require('../server/commands/githubLabels');
const assetSearch = require('../server/commands/assetSearch');
const hypervStatus = require('../server/commands/hypervStatus');
const processingStatus = require('../server/commands/processingStatus');
const globalVar = require('../server/commands/globalVar');
const help = require('../server/commands/help');

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
        // Invalid department returns { invalidDept } object, not null —
        // this signals to execute() that it should show an error with the
        // list of valid departments rather than submit a ticket.
        const m = ticketQuick.match('add foo ticket bar', 'add foo ticket bar', { ima: 'admin' });
        assert.ok(m);
        assert.equal(m.invalidDept, 'foo');
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

// ─── githubIssues ──────────────────────────────────────────────────────────

describe('githubIssues command', () => {
    const deps = { ima: 'admin' };
    it('matches "issues list"', () => {
        const m = githubIssues.match('issues list', 'issues list', deps);
        assert.ok(m);
        assert.equal(m.action, 'list');
    });
    it('matches "issues"', () => {
        const m = githubIssues.match('issues', 'issues', deps);
        assert.ok(m);
        assert.equal(m.action, 'list');
    });
    it('matches "issues show 42"', () => {
        const m = githubIssues.match('issues show 42', 'issues show 42', deps);
        assert.ok(m);
        assert.equal(m.action, 'show');
        assert.equal(m.id, 42);
    });
    it('matches "show issue 5"', () => {
        const m = githubIssues.match('show issue 5', 'show issue 5', deps);
        assert.ok(m);
        assert.equal(m.action, 'show');
        assert.equal(m.id, 5);
    });
    it('matches "issue close 10 fixed it"', () => {
        const m = githubIssues.match('issue close 10 fixed it', 'issue close 10 fixed it', deps);
        assert.ok(m);
        assert.equal(m.action, 'close');
        assert.equal(m.id, 10);
        assert.equal(m.comment, 'fixed it');
    });
    it('matches "issues close 7" without comment', () => {
        const m = githubIssues.match('issues close 7', 'issues close 7', deps);
        assert.ok(m);
        assert.equal(m.action, 'close');
        assert.equal(m.comment, '');
    });
    it('matches "issues comment 3 looks good"', () => {
        const m = githubIssues.match('issues comment 3 looks good', 'issues comment 3 looks good', deps);
        assert.ok(m);
        assert.equal(m.action, 'comment');
        assert.equal(m.id, 3);
        assert.equal(m.comment, 'looks good');
    });
    it('matches "issues create Fix the thing"', () => {
        const m = githubIssues.match('issues create Fix the thing', 'issues create fix the thing', deps);
        assert.ok(m);
        assert.equal(m.action, 'create');
        assert.equal(m.raw, 'Fix the thing');
    });
    it('matches "github help"', () => {
        const m = githubIssues.match('github help', 'github help', deps);
        assert.ok(m);
        assert.equal(m.action, 'help');
    });
    it('matches "gh help"', () => {
        const m = githubIssues.match('gh help', 'gh help', deps);
        assert.ok(m);
        assert.equal(m.action, 'help');
    });
    it('does not match for non-admin', () => {
        assert.equal(githubIssues.match('issues list', 'issues list', { ima: 'user' }), null);
    });
});

// ─── githubLabels ──────────────────────────────────────────────────────────

describe('githubLabels command', () => {
    const deps = { ima: 'admin' };
    it('matches "labels list"', () => {
        const m = githubLabels.match('labels list', 'labels list', deps);
        assert.ok(m);
        assert.equal(m.action, 'list');
    });
    it('matches "labels"', () => {
        const m = githubLabels.match('labels', 'labels', deps);
        assert.ok(m);
        assert.equal(m.action, 'list');
    });
    it('matches "label create bug ff0000"', () => {
        const m = githubLabels.match('label create bug ff0000', 'label create bug ff0000', deps);
        assert.ok(m);
        assert.equal(m.action, 'create');
        assert.equal(m.name, 'bug');
        assert.equal(m.color, 'ff0000');
    });
    it('matches "label create bug #00ff00 Description here"', () => {
        const m = githubLabels.match('label create bug #00ff00 Description here', 'label create bug #00ff00 description here', deps);
        assert.ok(m);
        assert.equal(m.color, '00ff00');
        assert.equal(m.description, 'Description here');
    });
    it('matches "label update bug fix 0000ff"', () => {
        const m = githubLabels.match('label update bug fix 0000ff', 'label update bug fix 0000ff', deps);
        assert.ok(m);
        assert.equal(m.action, 'update');
        assert.equal(m.name, 'bug');
        assert.equal(m.newName, 'fix');
        assert.equal(m.color, '0000ff');
    });
    it('matches "label add 5 enhancement"', () => {
        const m = githubLabels.match('label add 5 enhancement', 'label add 5 enhancement', deps);
        assert.ok(m);
        assert.equal(m.action, 'add');
        assert.equal(m.issueId, 5);
        assert.equal(m.label, 'enhancement');
    });
    it('matches "label remove 3 bug"', () => {
        const m = githubLabels.match('label remove 3 bug', 'label remove 3 bug', deps);
        assert.ok(m);
        assert.equal(m.action, 'remove');
        assert.equal(m.issueId, 3);
        assert.equal(m.label, 'bug');
    });
    it('does not match for non-admin', () => {
        assert.equal(githubLabels.match('labels list', 'labels list', { ima: 'user' }), null);
    });
});

// ─── assetSearch ───────────────────────────────────────────────────────────

describe('assetSearch command', () => {
    const deps = { ima: 'admin' };
    it('matches "search asset server1"', () => {
        const m = assetSearch.match('search asset server1', 'search asset server1', deps);
        assert.ok(m);
        assert.equal(m.query, 'server1');
    });
    it('matches "find asset 123"', () => {
        const m = assetSearch.match('find asset 123', 'find asset 123', deps);
        assert.ok(m);
        assert.equal(m.query, '123');
    });
    it('matches "lookup asset web-host.prod"', () => {
        const m = assetSearch.match('lookup asset web-host.prod', 'lookup asset web-host.prod', deps);
        assert.ok(m);
        assert.equal(m.query, 'web-host.prod');
    });
    it('does not match for non-admin', () => {
        assert.equal(assetSearch.match('search asset foo', 'search asset foo', { ima: 'user' }), null);
    });
    it('does not match without asset keyword', () => {
        assert.equal(assetSearch.match('search server1', 'search server1', deps), null);
    });
});

// ─── hypervStatus ──────────────────────────────────────────────────────────

describe('hypervStatus command', () => {
    it('matches "hyperv status"', () => {
        assert.ok(hypervStatus.match('hyperv status', 'hyperv status', { ima: 'admin' }));
    });
    it('matches case-insensitively', () => {
        assert.ok(hypervStatus.match('HyperV Status', 'hyperv status', { ima: 'admin' }));
    });
    it('does not match for non-admin', () => {
        assert.equal(hypervStatus.match('hyperv status', 'hyperv status', { ima: 'user' }), null);
    });
    it('does not match partial text', () => {
        assert.equal(hypervStatus.match('hyperv', 'hyperv', { ima: 'admin' }), null);
    });
});

// ─── processingStatus ──────────────────────────────────────────────────────

describe('processingStatus command', () => {
    it('matches "processing status"', () => {
        assert.ok(processingStatus.match('processing status', 'processing status', { ima: 'admin' }));
    });
    it('does not match for non-admin', () => {
        assert.equal(processingStatus.match('processing status', 'processing status', { ima: 'user' }), null);
    });
    it('does not match partial text', () => {
        assert.equal(processingStatus.match('processing', 'processing', { ima: 'admin' }), null);
    });
});

// ─── globalVar ─────────────────────────────────────────────────────────────

describe('globalVar command', () => {
    const deps = { ima: 'admin' };
    it('matches "get global myvar"', () => {
        const m = globalVar.match('get global myvar', 'get global myvar', deps);
        assert.ok(m);
        assert.equal(m.action, 'get');
        assert.equal(m.varName, 'myvar');
    });
    it('matches "set global myvar hello world"', () => {
        const m = globalVar.match('set global myvar hello world', 'set global myvar hello world', deps);
        assert.ok(m);
        assert.equal(m.action, 'set');
        assert.equal(m.varName, 'myvar');
        assert.equal(m.value, 'hello world');
    });
    it('does not match for non-admin', () => {
        assert.equal(globalVar.match('get global foo', 'get global foo', { ima: 'user' }), null);
    });
    it('does not match bare "get global"', () => {
        assert.equal(globalVar.match('get global', 'get global', deps), null);
    });
});

// ─── help ──────────────────────────────────────────────────────────────────

describe('help command', () => {
    it('matches "help"', () => {
        assert.ok(help.match('help', 'help'));
    });
    it('does not match "help me"', () => {
        assert.equal(help.match('help me', 'help me'), null);
    });
    it('does not match "github help"', () => {
        assert.equal(help.match('github help', 'github help'), null);
    });
});
