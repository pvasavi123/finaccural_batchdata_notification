const IOAuthTokenRepository = require('../../../core/oauth/interfaces/IOAuthTokenRepository');
const { XeroToken } = require('../../../core/database');

class XeroTokenRepository extends IOAuthTokenRepository {
    async getToken(tenantId) {
        const token = await XeroToken.findOne({ where: { tenant_id: tenantId } });
        if (!token) return null;

        // Calculate expires_at dynamically. If `expires_in` is an absolute timestamp (large number), use it directly.
        // Otherwise, fall back to calculating from `updatedAt`.
        const isOldFormat = token.expires_in < 1000000000;
        const expiresAt = isOldFormat 
            ? new Date(token.updatedAt.getTime() + token.expires_in * 1000)
            : new Date(token.expires_in * 1000);

        return {
            accessToken: token.access_token,
            refreshToken: token.refresh_token,
            expiresAt: expiresAt,
            rawToken: token
        };
    }

    async saveToken(tenantId, { accessToken, refreshToken, expiresAt }) {
        const expiresIn = Math.floor(expiresAt.getTime() / 1000); // Store absolute UNIX timestamp
        const { Op } = require('sequelize');

        // Note: Xero supports multi-tenancy. If multiple tenants are connected under the same
        // session/refresh flow, we must update the token for ALL records sharing the old refresh token.
        const currentToken = await XeroToken.findOne({ where: { tenant_id: tenantId } });
        if (currentToken) {
            // Find all tenants that share this token set
            await XeroToken.update({
                access_token: accessToken,
                refresh_token: refreshToken,
                expires_in: expiresIn
            }, {
                where: { refresh_token: currentToken.refresh_token }
            });

            // Recovering from 'Disconnected' re-activates the connection.
            // A connection that hasn't had its first successful Master Data
            // Pull yet stays 'Not Synced' even though its token was just
            // refreshed — refreshing a token isn't the same as syncing data.
            await XeroToken.update({
                status: 'Active'
            }, {
                where: { refresh_token: refreshToken, status: { [Op.ne]: 'Not Synced' } }
            });
        } else {
            // Fallback for single tenant update
            await XeroToken.update({
                access_token: accessToken,
                refresh_token: refreshToken,
                expires_in: expiresIn
            }, {
                where: { tenant_id: tenantId }
            });

            await XeroToken.update({
                status: 'Active'
            }, {
                where: { tenant_id: tenantId, status: { [Op.ne]: 'Not Synced' } }
            });
        }
    }

    /**
     * Marks the connection(s) as needing reconnection and wipes the now-
     * unusable credentials — access token, refresh token, and any stored
     * provider session — rather than just flipping the status flag.
     * `access_token`/`refresh_token`/`expires_in` are NOT NULL columns, so
     * they're cleared to an empty string/0 (falsy, same practical effect as
     * null, no schema change needed).
     */
    async markDisconnected(tenantId) {
        const clearedFields = {
            status: 'Disconnected',
            access_token: '',
            refresh_token: '',
            expires_in: 0,
            session_info: null
        };

        const currentToken = await XeroToken.findOne({ where: { tenant_id: tenantId } });
        if (currentToken) {
            // Disconnect all tenants associated with this credential set —
            // they share the same (now-revoked) refresh token.
            await XeroToken.update(clearedFields, {
                where: { refresh_token: currentToken.refresh_token }
            });
        } else {
            await XeroToken.update(clearedFields, { where: { tenant_id: tenantId } });
        }
    }
}

module.exports = XeroTokenRepository;
