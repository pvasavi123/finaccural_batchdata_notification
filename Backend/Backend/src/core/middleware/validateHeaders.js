'use strict';

const { ValidationError } = require('../errors/AppError');

/**
 * validateContentType
 * ----------------------------------------------------------------
 * Header Validation.
 *
 * Every route in this API that accepts a JSON body (POST/PUT/PATCH)
 * expects `Content-Type: application/json`. Previously the only header
 * FinAccrual actually validated was `Authorization: Bearer <jwt>`
 * (in auth.middleware.js) — the Content-Type header itself was never
 * checked, so a malformed or missing Content-Type would either silently
 * pass through as an empty `req.body` (express.json() just skips
 * parsing) or fail deep inside a controller with a confusing
 * "Cannot read properties of undefined" error instead of a clean 400.
 *
 * Mounted globally, BEFORE express.json() — it only inspects raw
 * headers (never req.body), so there's no need to let body-parsing run
 * first, and rejecting early avoids doing that parsing work at all for
 * a request that's going to be rejected anyway. GET/DELETE/HEAD/OPTIONS
 * requests (no body expected) and requests with an empty body are
 * skipped — this only enforces the header when a body is actually
 * expected to be present.
 * ----------------------------------------------------------------
 */
const METHODS_WITH_BODY = new Set(['POST', 'PUT', 'PATCH']);

function validateContentType(req, res, next) {
    if (!METHODS_WITH_BODY.has(req.method)) {
        return next();
    }

    // No body at all (e.g. POST /api/auth/logout with an empty payload) —
    // nothing to validate.
    const contentLength = req.headers['content-length'];
    if (!contentLength || contentLength === '0') {
        return next();
    }

    const contentType = req.headers['content-type'];

    if (!contentType) {
        return next(new ValidationError(
            'Missing Content-Type header.',
            `${req.method} ${req.originalUrl} sent a body without a Content-Type header.`
        ));
    }

    if (!contentType.toLowerCase().includes('application/json')) {
        return next(new ValidationError(
            'Unsupported Content-Type. Expected application/json.',
            `${req.method} ${req.originalUrl} sent Content-Type: ${contentType}`
        ));
    }

    next();
}

module.exports = { validateContentType };
