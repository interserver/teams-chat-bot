const { describe, it } = require('node:test');
const assert = require('node:assert/strict');

const {
    parseRepoPatterns,
    repoMatchesPattern,
    decideAnnounceRedirect
} = require('../server/queue/notificationConsumer');

describe('parseRepoPatterns', () => {
    it('returns empty array for undefined/empty input', () => {
        assert.deepEqual(parseRepoPatterns(undefined), []);
        assert.deepEqual(parseRepoPatterns(''), []);
        assert.deepEqual(parseRepoPatterns('   '), []);
    });

    it('splits comma-separated patterns and trims whitespace', () => {
        assert.deepEqual(
            parseRepoPatterns('detain/*, owner/repo , detain/foo'),
            ['detain/*', 'owner/repo', 'detain/foo']
        );
    });

    it('drops empty segments from extra commas', () => {
        assert.deepEqual(
            parseRepoPatterns('detain/*,,owner/repo,'),
            ['detain/*', 'owner/repo']
        );
    });
});

describe('repoMatchesPattern', () => {
    it('matches exact owner/repo', () => {
        assert.equal(repoMatchesPattern('detain/myadmin', 'detain/myadmin'), true);
        assert.equal(repoMatchesPattern('detain/myadmin', 'detain/other'), false);
    });

    it('matches wildcard owner/*', () => {
        assert.equal(repoMatchesPattern('detain/myadmin', 'detain/*'), true);
        assert.equal(repoMatchesPattern('detain/sugarcraft', 'detain/*'), true);
        assert.equal(repoMatchesPattern('interserver/foo', 'detain/*'), false);
    });

    it('does not allow cross-org false positives', () => {
        // "detainx/foo" must not match "detain/*"
        assert.equal(repoMatchesPattern('detainx/foo', 'detain/*'), false);
    });

    it('returns false for empty repo or pattern', () => {
        assert.equal(repoMatchesPattern('', 'detain/*'), false);
        assert.equal(repoMatchesPattern('detain/foo', ''), false);
        assert.equal(repoMatchesPattern('', ''), false);
    });
});

describe('decideAnnounceRedirect', () => {
    it('returns no-redirect when announce list is empty', () => {
        const out = decideAnnounceRedirect('detain/myadmin', '', '');
        assert.deepEqual(out, { redirect: false });
    });

    it('redirects on wildcard announce match with no exclude', () => {
        const out = decideAnnounceRedirect('detain/myadmin', 'detain/*', '');
        assert.equal(out.redirect, true);
        assert.deepEqual(out.matched, ['detain/*']);
    });

    it('redirects on exact announce match', () => {
        const out = decideAnnounceRedirect('detain/myadmin', 'detain/myadmin', '');
        assert.equal(out.redirect, true);
        assert.deepEqual(out.matched, ['detain/myadmin']);
    });

    it('does NOT redirect when announce matches but exclude also matches (exact)', () => {
        const out = decideAnnounceRedirect(
            'detain/myadmin',
            'detain/*',
            'detain/myadmin'
        );
        assert.equal(out.redirect, false);
        assert.equal(out.excluded, true);
        assert.deepEqual(out.matched, ['detain/*']);
        assert.deepEqual(out.excludedBy, ['detain/myadmin']);
    });

    it('does NOT redirect when announce matches and exclude wildcard matches', () => {
        const out = decideAnnounceRedirect(
            'detain/sugarcraft',
            'detain/*',
            'detain/sugarcraft, detain/myadmin'
        );
        assert.equal(out.redirect, false);
        assert.equal(out.excluded, true);
    });

    it('still redirects when repo matches announce but NOT exclude', () => {
        const out = decideAnnounceRedirect(
            'detain/other-repo',
            'detain/*',
            'detain/myadmin, detain/sugarcraft'
        );
        assert.equal(out.redirect, true);
        assert.deepEqual(out.matched, ['detain/*']);
    });

    it('exclude has no effect when announce does not match', () => {
        const out = decideAnnounceRedirect(
            'interserver/foo',
            'detain/*',
            'interserver/foo'
        );
        assert.deepEqual(out, { redirect: false });
    });

    it('returns no-redirect for empty repo', () => {
        const out = decideAnnounceRedirect('', 'detain/*', '');
        assert.deepEqual(out, { redirect: false });
    });

    it('handles multiple announce patterns', () => {
        const out = decideAnnounceRedirect(
            'interserver/foo',
            'detain/*, interserver/*',
            ''
        );
        assert.equal(out.redirect, true);
        assert.deepEqual(out.matched, ['interserver/*']);
    });

    it('exclude wildcard exempts an entire sub-org', () => {
        // hypothetical: announce all of detain/* but exempt anything matching detain/*
        // (degenerate but documents behavior: exclude wins)
        const out = decideAnnounceRedirect(
            'detain/anything',
            'detain/*',
            'detain/*'
        );
        assert.equal(out.redirect, false);
        assert.equal(out.excluded, true);
    });
});
