'use strict';

const jwt    = require('jsonwebtoken');
const config = require('../../core/config');

/**
 * JwtService
 * ----------------------------------------------------------------
 * Centralised JWT utilities used by AuthService and AuthMiddleware.
 * Keeping JWT logic here means the secret and expiry are defined in
 * exactly one place.
 * ----------------------------------------------------------------
 */
class JwtService {

    /**
     * Sign a short-lived access JWT (15 minutes).
     * @param {{ userId: string, email: string, role: string }} payload
     * @returns {string} signed JWT
     */
    static generateToken(payload) {
        return jwt.sign(payload, config.JWT_SECRET, { expiresIn: '30m' });
    }

    /** Alias kept for readability in callers. */
    static generateAccessToken(payload) {
        return JwtService.generateToken(payload);
    }

    /**
     * Generate a cryptographically-random opaque refresh token (64 hex chars).
     * This is stored in the DB and looked up on /api/auth/refresh — it is
     * intentionally NOT a JWT so it cannot be decoded client-side.
     * @returns {string}
     */
    static generateRefreshToken() {
        const crypto = require('crypto');
        return crypto.randomBytes(64).toString('hex');
    }

    /**
     * Verify and decode an access JWT.
     * Throws if the token is invalid or expired.
     * @param {string} token
     * @returns {object} decoded payload
     */
    static verifyToken(token) {
        return jwt.verify(token, config.JWT_SECRET);
    }

    /** Alias kept for readability in callers. */
    static verifyAccessToken(token) {
        return JwtService.verifyToken(token);
    }
}

module.exports = JwtService;
