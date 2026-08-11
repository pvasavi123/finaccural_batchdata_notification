'use strict';

/**
 * sanitizeInput
 * ----------------------------------------------------------------
 * Security Validation.
 *
 * FinAccrual previously did no input sanitization at all — a string
 * field (name, companyName, email, etc.) was passed straight from
 * `req.body`/`req.query`/`req.params` through to Sequelize/QuickBooks-
 * Xero calls and, eventually, into an Excel cell. Sequelize's
 * parameterized queries already prevent classic SQL injection, but
 * nothing stripped stray control characters or `<script>`-style HTML
 * that a user could type into a "company name" field and have it
 * echoed back into a browser (e.g. in a UI list) or into an exported
 * XLSX cell later opened in Excel.
 *
 * This runs a lightweight, dependency-free pass, deliberately NOT a
 * full HTML sanitizer library (that would be overkill for a JSON API
 * that returns data to a React app, which already escapes on render) —
 * it just:
 *   1. Trims leading/trailing whitespace on every string field.
 *   2. Strips raw ASCII control characters (0x00-0x1F, 0x7F) that have
 *      no legitimate place in a name/email/company field.
 *   3. Strips `<script>...</script>` blocks and inline `on*=` handler
 *      attributes as a defense-in-depth measure against stored XSS,
 *      in case a value is ever rendered as raw HTML somewhere downstream.
 *
 * Mounted globally, immediately after express.json() (it needs
 * req.body to already be a parsed object), before the routes.
 * ----------------------------------------------------------------
 */

const CONTROL_CHARS = /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g; // eslint-disable-line no-control-regex
const SCRIPT_TAG    = /<script[\s\S]*?<\/script>/gi;
const ON_ATTR       = /\son\w+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi;

function sanitizeString(str) {
    return str
        .replace(SCRIPT_TAG, '')
        .replace(ON_ATTR, '')
        .replace(CONTROL_CHARS, '')
        .trim();
}

function sanitizeValue(value, depth = 0) {
    if (depth > 10) return value; // guard against pathological/circular input

    if (typeof value === 'string') {
        return sanitizeString(value);
    }
    if (Array.isArray(value)) {
        return value.map((v) => sanitizeValue(v, depth + 1));
    }
    if (value && typeof value === 'object') {
        const out = {};
        for (const key of Object.keys(value)) {
            out[key] = sanitizeValue(value[key], depth + 1);
        }
        return out;
    }
    return value;
}

function sanitizeInput(req, res, next) {
    // req.body is a plain, reassignable object (set by express.json()) —
    // safe to overwrite in place.
    if (req.body && typeof req.body === 'object') {
        req.body = sanitizeValue(req.body);
    }

    // req.params is a plain per-request object — safe to mutate in place.
    if (req.params && typeof req.params === 'object') {
        for (const key of Object.keys(req.params)) {
            req.params[key] = sanitizeValue(req.params[key]);
        }
    }

    // NOTE: req.query is intentionally NOT sanitized here. In Express 5,
    // req.query is a computed getter re-derived from the raw URL on every
    // access, so both reassigning it (`req.query = ...`) and mutating the
    // object it returns are silently lost by the next access — there is
    // no supported way to persist a change to it in place. Query params
    // in this API are narrow (tier/platform/companyId/plan) and are
    // already constrained to fixed enums or DB-lookup values by the Joi
    // schemas in core/validation/schemas.js, so the XSS surface there is
    // covered by validation rather than sanitization.
    next();
}

module.exports = { sanitizeInput, sanitizeString };
