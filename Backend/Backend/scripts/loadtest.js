'use strict';

/**
 * Load / Performance Test
 * ----------------------------------------------------------------
 * Performance Validation.
 *
 * FinAccrual previously had NO load/stress testing of any kind — this
 * is the first one. It uses autocannon (installed as a devDependency)
 * to hammer a couple of representative endpoints and report throughput
 * (req/sec), latency percentiles, and error counts.
 *
 * This intentionally does NOT hit real QuickBooks/Xero endpoints or
 * require a logged-in JWT for every path — it targets:
 *   1. GET /favicon.ico       — cheapest possible route, measures pure
 *                                Express + middleware-chain overhead
 *                                (helmet, cors, sanitizeInput,
 *                                responseTime, rate limiters, etc.)
 *   2. POST /api/auth/login   — the most-hit real endpoint, with
 *                                deliberately wrong credentials, to
 *                                measure the bcrypt-heavy path AND to
 *                                confirm the authLimiter (429) actually
 *                                kicks in under load.
 *
 * Usage:
 *   node scripts/loadtest.js                    # against http://localhost:8000
 *   BASE_URL=http://localhost:5000 node scripts/loadtest.js
 * ----------------------------------------------------------------
 */

const autocannon = require('autocannon');

const BASE_URL = process.env.BASE_URL || 'http://localhost:8000';

function run(opts) {
    return new Promise((resolve, reject) => {
        const instance = autocannon(opts, (err, result) => {
            if (err) return reject(err);
            resolve(result);
        });
        autocannon.track(instance, { renderProgressBar: true });
    });
}

function printSummary(label, result) {
    console.log(`\n=== ${label} ===`);
    console.log(`Requests:        ${result.requests.total}`);
    console.log(`Throughput:      ${result.requests.average} req/sec (avg)`);
    console.log(`Latency (avg):   ${result.latency.average} ms`);
    console.log(`Latency (p99):   ${result.latency.p99} ms`);
    console.log(`2xx responses:   ${result['2xx']}`);
    console.log(`Non-2xx/errors:  ${result.errors + (result.non2xx || 0)}`);
}

async function main() {
    console.log(`Target: ${BASE_URL}`);
    console.log('This expects the server to already be running (npm start / npm run dev).');

    const smoke = await run({
        url: `${BASE_URL}/favicon.ico`,
        connections: 20,
        duration:    10
    });
    printSummary('GET /favicon.ico (baseline overhead)', smoke);

    const login = await run({
        url: `${BASE_URL}/api/auth/login`,
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ email: 'loadtest@example.com', password: 'wrong-password' }),
        connections: 10,
        duration:    10
    });
    printSummary('POST /api/auth/login (expect 401s, then 429s once authLimiter trips)', login);

    console.log('\nDone. Compare Latency p99 against the SLOW_REQUEST_MS threshold ' +
        '(core/middleware/responseTime.js, currently 2000ms).');
}

main().catch((err) => {
    console.error('Load test failed:', err.message);
    console.error('Is the server running? Try `npm run dev` in another terminal first.');
    process.exit(1);
});
