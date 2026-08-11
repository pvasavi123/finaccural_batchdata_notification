'use strict';

const { ValidationError } = require('../errors/AppError');

/**
 * validate(schema, source)
 * ----------------------------------------------------------------
 * Data Type Validation.
 *
 * Generic Joi-schema validation middleware factory. Previously
 * FinAccrual only had hand-rolled regex checks (auth.validation.js) for
 * signup/login and no schema validation at all for query params or
 * other route bodies — a wrong data TYPE (e.g. `tier: 123` instead of
 * `tier: "pro"`, or an extra unexpected field) would sail straight
 * through to the service/DB layer instead of being rejected at the
 * edge with a clear 400.
 *
 * Usage:
 *   router.post('/signup', validate(schemas.signup), controller.signup);
 *   router.get('/pull-master-data', validate(schemas.pullMasterDataQuery, 'query'), controller.pullMasterData);
 *
 * On success, `req.body` is replaced with the *validated + coerced*
 * value (e.g. Joi will cast a numeric string to the declared type).
 * `req.query`/`req.params` are validated the same way but are NOT
 * reassigned — Express 5 exposes `req.query` as a computed getter (it
 * re-derives the object from the raw URL on every access), so
 * `req.query = value` is silently a no-op. Controllers already read
 * query params as raw strings (e.g. `req.query.tier.toLowerCase()`), so
 * this only needs to reject invalid input at the edge, not rewrite it.
 * ----------------------------------------------------------------
 * @param {import('joi').ObjectSchema} schema
 * @param {'body'|'query'|'params'} [source='body']
 */
function validate(schema, source = 'body') {
    return (req, res, next) => {
        const { error, value } = schema.validate(req[source], {
            abortEarly:      false, // collect every failing field, not just the first
            stripUnknown:    true,  // drop fields the schema doesn't declare
            convert:         true   // coerce e.g. "3" -> 3 where the schema expects a number
        });

        if (error) {
            const message = error.details.map((d) => d.message).join('; ');
            return next(new ValidationError(message, error.details));
        }

        if (source === 'body') {
            req.body = value;
        }
        next();
    };
}

module.exports = { validate };
