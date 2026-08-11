'use strict';

const logger = require('../logger');
const {
    AppError,
    ConnectionRefusedError,
    SessionExpiredError,
    ValidationError
} = require('../errors/AppError');

// Node/axios/mysql2 error codes that mean "couldn't reach a dependency"
// (DB down, upstream API host down/refusing connections, DNS failure, etc.)
const NETWORK_ERROR_CODES = new Set([
    'ECONNREFUSED',
    'ENOTFOUND',
    'ETIMEDOUT',
    'EAI_AGAIN',
    'ECONNRESET',
    'EHOSTUNREACH',
    'ENETUNREACH'
]);

const DB_CONNECTION_ERROR_NAMES = new Set([
    'SequelizeConnectionError',
    'SequelizeConnectionRefusedError',
    'SequelizeHostNotReachableError',
    'SequelizeConnectionTimedOutError',
    'SequelizeAccessDeniedError'
]);

// Database Validation: a Sequelize model `validate: {...}` rule (see
// user.model.js, quickbooks/model.js, xero/model.js) or a unique-index
// violation (email already registered) failed. Without this branch these
// fell through to the generic 500 ERR_INTERNAL below — which is wrong,
// since a rejected write because the DATA was invalid is a client error
// (400), not a server failure.
const DB_VALIDATION_ERROR_NAMES = new Set([
    'SequelizeValidationError',
    'SequelizeUniqueConstraintError'
]);

/**
 * Normalize ANY thrown value into an AppError so the response shape is
 * always consistent, even for errors we didn't anticipate (raw Node
 * errors, third-party library errors, typos, etc.). This is the single
 * choke point that guarantees users never see a raw stack trace, SQL
 * error, or JWT library message.
 */
function normalize(err) {
    if (err instanceof AppError) return err;

    if (err && NETWORK_ERROR_CODES.has(err.code)) {
        return new ConnectionRefusedError(err.message);
    }

    if (err && DB_CONNECTION_ERROR_NAMES.has(err.name)) {
        return new ConnectionRefusedError(err.message);
    }

    if (err && DB_VALIDATION_ERROR_NAMES.has(err.name)) {
        // err.errors is Sequelize's array of { message, path, ... } —
        // collapse it into one readable string instead of exposing the
        // raw Sequelize error shape to the client.
        const detail = Array.isArray(err.errors) && err.errors.length
            ? err.errors.map((e) => e.message).join('; ')
            : err.message;
        return new ValidationError(detail, detail);
    }

    // Defence in depth: a raw jsonwebtoken error that slipped past
    // auth.middleware.js (e.g. thrown from somewhere else in the app).
    if (err && (err.name === 'TokenExpiredError' || err.name === 'JsonWebTokenError' || err.name === 'NotBeforeError')) {
        return new SessionExpiredError(err.message);
    }

    // Anything else is unexpected — treat as an internal error and never
    // leak the raw message/stack to the client.
    return new AppError(
        'Something went wrong on our end. Please try again later.',
        500,
        'ERR_INTERNAL',
        (err && err.message) || 'Unknown error'
    );
}

/**
 * Centralized Express error-handling middleware. Mount this LAST, after
 * all routes. Every controller should funnel failures here via
 * `next(err)` (or by throwing inside a handler wrapped with
 * `asyncHandler`) instead of formatting its own response.
 */
// eslint-disable-next-line no-unused-vars
function errorHandler(err, req, res, next) {
    const appError = normalize(err);

    // Full technical detail (stack included) goes to server logs ONLY.
    logger.error(
        `[${req.method} ${req.originalUrl}] ${appError.code} (${appError.statusCode}): ${err && err.message}`,
        err && err.stack ? err.stack : err
    );

    return res.status(appError.statusCode).json({
        success: false,
        code:    appError.code,
        message: appError.message,
        details: appError.details
    });
}

/** Mount right before errorHandler to turn unknown routes into a standard 404. */
function notFoundHandler(req, res, next) {
    next(new AppError(
        'The requested resource was not found.',
        404,
        'ERR_NOT_FOUND',
        `Route ${req.method} ${req.originalUrl} not found.`
    ));
}

module.exports = { errorHandler, notFoundHandler, normalize };
