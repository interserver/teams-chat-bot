// Shared retry helper for outbound Bot Framework calls.
//
// Replaces the inline 2-retry linear-backoff loops that lived in
// botController.js / msgController.js / dailyRecapController.js with a
// single implementation that adds exponential backoff and jitter. On
// auth failures we still re-trust the serviceUrl, matching the original
// behaviour.

const { MicrosoftAppCredentials } = require('botframework-connector');

const TRANSIENT_RE = /ECONNRESET|ETIMEDOUT|ENOTFOUND|socket hang up|EAI_AGAIN/i;
const AUTH_ERROR_RE = /authorization has been denied|401|unauthorized/i;
const RATE_LIMIT_RE = /429|too many requests|throttl/i;

const DEFAULT_OPTS = {
    maxRetries: 3,
    baseDelayMs: 750,
    maxDelayMs: 8000,
    onAuthFailure: null   // function(serviceUrl) — called once per auth retry
};

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function classify(err) {
    const msg = (err && err.message) || '';
    return {
        transient: TRANSIENT_RE.test(msg),
        auth: AUTH_ERROR_RE.test(msg),
        rateLimit: RATE_LIMIT_RE.test(msg)
    };
}

function backoffDelay(attempt, opts) {
    const exp = opts.baseDelayMs * Math.pow(2, attempt - 1);
    const capped = Math.min(exp, opts.maxDelayMs);
    const jitter = capped * (0.7 + Math.random() * 0.6);
    return Math.round(jitter);
}

/**
 * Run `fn` with retries on transient/auth/rate-limit failures.
 *
 * @param {Function} fn  async () => any. Should throw on failure.
 * @param {Object}   options { maxRetries, baseDelayMs, maxDelayMs, serviceUrl?, label? }
 */
async function runWithRetry(fn, options = {}) {
    const opts = { ...DEFAULT_OPTS, ...options };
    const label = opts.label || 'send';
    let lastErr;
    for (let attempt = 1; attempt <= opts.maxRetries; attempt++) {
        try {
            return await fn(attempt);
        } catch (err) {
            lastErr = err;
            const cls = classify(err);
            const retriable = cls.transient || cls.auth || cls.rateLimit;
            if (!retriable || attempt === opts.maxRetries) {
                throw err;
            }
            if (cls.auth && opts.serviceUrl) {
                try {
                    MicrosoftAppCredentials.trustServiceUrl(opts.serviceUrl);
                } catch (_) {
                    // best-effort; we'll retry regardless
                }
            }
            const delay = backoffDelay(attempt, opts);
            console.warn(`[${ label }] attempt ${ attempt }/${ opts.maxRetries } failed (${ err.message }); retrying in ${ delay }ms`);
            await sleep(delay);
        }
    }
    throw lastErr;
}

module.exports = { runWithRetry, classify, sleep, backoffDelay };
