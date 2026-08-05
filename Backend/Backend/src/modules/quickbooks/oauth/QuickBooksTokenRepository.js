const IOAuthTokenRepository = require('../../../core/oauth/interfaces/IOAuthTokenRepository');
const { QuickBooksToken } = require('../../../core/database');

class QuickBooksTokenRepository extends IOAuthTokenRepository {
    async getToken(realmId) {
        const token = await QuickBooksToken.findOne({ where: { realm_id: realmId } });
        if (!token) return null;

        // If expires_in is small (e.g., 3600), it's from the old format. 
        // We calculate expiresAt relative to updated_at. If it's a large timestamp, we use it directly.
        const isOldFormat = token.expires_in < 1000000000;
        const expiresAt = isOldFormat 
            ? new Date(token.updated_at.getTime() + token.expires_in * 1000)
            : new Date(token.expires_in * 1000);

        // Refresh token expiry: QuickBooks refresh tokens live for 100 days.
        let refreshTokenExpiresAt = null;
        if (token.x_refresh_token_expires_in && token.x_refresh_token_expires_in > 0) {
            const isRefreshOldFormat = token.x_refresh_token_expires_in < 1000000000;
            refreshTokenExpiresAt = isRefreshOldFormat
                ? new Date(token.updated_at.getTime() + token.x_refresh_token_expires_in * 1000)
                : new Date(token.x_refresh_token_expires_in * 1000);
        }

        return {
            accessToken:           token.access_token,
            refreshToken:          token.refresh_token,
            expiresAt:             expiresAt,
            refreshTokenExpiresAt: refreshTokenExpiresAt,
            rawToken:              token
        };
    }

    async saveToken(realmId, { accessToken, refreshToken, expiresAt, refreshTokenExpiresIn }) {
        const updatePayload = {
            access_token:  accessToken,
            refresh_token: refreshToken,
            expires_in:    Math.floor(expiresAt.getTime() / 1000) // Store absolute UNIX timestamp
        };

        // Persist the refresh token lifetime when provided by the OAuth response
        if (refreshTokenExpiresIn != null && refreshTokenExpiresIn > 0) {
            updatePayload.x_refresh_token_expires_in = Math.floor(Date.now() / 1000) + refreshTokenExpiresIn;
        }

        await QuickBooksToken.update(updatePayload, {
            where: { realm_id: realmId }
        });

        // Recovering from 'Disconnected' re-activates the connection once a
        // refresh succeeds. But a connection that hasn't had its first
        // successful Master Data Pull yet must stay 'Not Synced' — merely
        // refreshing an access token isn't a sync, so it's excluded here.
        const { Op } = require('sequelize');
        await QuickBooksToken.update(
            { status: 'Active' },
            { where: { realm_id: realmId, status: { [Op.ne]: 'Not Synced' } } }
        );
    }

    /**
     * Marks the connection as needing reconnection and wipes the now-unusable
     * credentials — access token, refresh token, and any stored provider
     * session — rather than just flipping the status flag. `access_token`/
     * `refresh_token` are NOT NULL columns, so they're cleared to an empty
     * string (falsy, same practical effect as null, no schema change needed).
     */
    async markDisconnected(realmId) {
        await QuickBooksToken.update({
            status: 'Disconnected',
            access_token: '',
            refresh_token: '',
            expires_in: 0,
            x_refresh_token_expires_in: 0,
            session_info: null
        }, {
            where: { realm_id: realmId }
        });
    }
}

module.exports = QuickBooksTokenRepository;
