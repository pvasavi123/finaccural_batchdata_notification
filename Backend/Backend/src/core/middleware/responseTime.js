'use strict';

const logger = require('../logger');

/**
 * responseTime
 * ----------------------------------------------------------------
 * Performance Validation.
 *
 * FinAccrual previously had no timing/threshold check on any request —
 * batching (batchDataLoader.js on the frontend, queryAll() paging on the
 * backend) was used to keep Excel writes and QuickBooks/Xero pulls
 * responsive, but nothing ever measured whether a request actually met
 * a response-time budget, and nothing flagged a request that didn't.
 *
 * This is intentionally simple: it does NOT replace real load/stress
 * testing (see scripts/loadtest.js for that) — it's an always-on,
 * zero-dependency guardrail that:
 *   1. Adds an `X-Response-Time` header to every response (ms).
 *   2. Logs a warning for any request that exceeds SLOW_REQUEST_MS, so
 *      slow QuickBooks/Xero pulls or Excel-export requests show up in
 *      the server logs instead of silently degrading.
 *
 * Mounted globally, first in the middleware chain (before body parsing)
 * so the timer covers the full request lifecycle.
 * ----------------------------------------------------------------
 */

// Any request slower than this is logged as a warning. 2s is generous —
// master-data pulls that fan out to multiple QuickBooks/Xero companies
// are the slowest legitimate operation in the app.
const SLOW_REQUEST_MS = 2000;

function responseTime(req, res, next) {
    const start = process.hrtime.bigint();

    // Headers can only be set BEFORE the response is flushed, so hook
    // res.end() (not the 'finish' event, which fires after headers are
    // already on the wire) to attach X-Response-Time just in time.
    const originalEnd = res.end;
    res.end = function patchedEnd(...args) {
        const durationMs = Number(process.hrtime.bigint() - start) / 1e6;
        if (!res.headersSent) {
            res.setHeader('X-Response-Time', `${durationMs.toFixed(1)}ms`);
        }
        return originalEnd.apply(res, args);
    };

    res.on('finish', () => {
        const durationMs = Number(process.hrtime.bigint() - start) / 1e6;
        if (durationMs > SLOW_REQUEST_MS) {
            logger.warn(
                `[SLOW REQUEST] ${req.method} ${req.originalUrl} took ${durationMs.toFixed(0)}ms ` +
                `(threshold: ${SLOW_REQUEST_MS}ms), status ${res.statusCode}`
            );
        }
    });

    next();
}

module.exports = { responseTime, SLOW_REQUEST_MS };
