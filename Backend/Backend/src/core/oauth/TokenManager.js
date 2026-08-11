const { OAuthTokenRevokedError, OAuthTokenRefreshError } = require('./errors/OAuthErrors');

class TokenManager {
    /**
     * @param {string} providerName - 'quickbooks' or 'xero' (for error logging)
     * @param {IOAuthClient} oauthClient - Client wrapping API logic
     * @param {IOAuthTokenRepository} tokenRepository - DB Adapter
     * @param {ILockManager} lockManager - Concurrency Lock Manager
     */
    constructor(providerName, oauthClient, tokenRepository, lockManager) {
        this.providerName = providerName;
        this.oauthClient = oauthClient;
        this.tokenRepository = tokenRepository;
        this.lockManager = lockManager;
    }

    /**
     * Check, auto-refresh, and return a valid access token.
     * @param {string} accountId - e.g. realmId or tenantId
     * @returns {Promise<string>} Valid Access Token
     */
    async getValidToken(accountId) {
        let tokenRecord = await this.tokenRepository.getToken(accountId);
        if (!tokenRecord) {
            throw new OAuthTokenRevokedError(
                `${this.providerName} is not connected for account ${accountId}.`,
                this.providerName
            );
        }

        // If access token is valid and expires more than 5 minutes from now, return it
        if (!this.isExpiringSoon(tokenRecord.expiresAt)) {
            return tokenRecord.accessToken;
        }

        // Token requires refresh. Acquire lock to prevent concurrent requests
        const releaseLock = await this.lockManager.acquire(accountId);
        try {
            // Re-check token record from DB inside lock in case another request refreshed it
            tokenRecord = await this.tokenRepository.getToken(accountId);
            if (!this.isExpiringSoon(tokenRecord.expiresAt)) {
                return tokenRecord.accessToken;
            }

            // ── Pre-flight: check if the refresh token itself has expired ──
            // If the repository provides refreshTokenExpiresAt (e.g. QuickBooks
            // x_refresh_token_expires_in), we can skip the network call entirely
            // and immediately mark the connection as disconnected.
            if (tokenRecord.refreshTokenExpiresAt &&
                new Date(tokenRecord.refreshTokenExpiresAt).getTime() <= Date.now()) {
                await this.tokenRepository.markDisconnected(accountId);
                throw new OAuthTokenRevokedError(
                    `${this.providerName} refresh token has expired for account ${accountId}. ` +
                    `User must re-authenticate.`,
                    this.providerName
                );
            }

            // Perform refresh call
            const newTokens = await this.oauthClient.refreshTokens(tokenRecord.refreshToken);

            // Calculate exact expiry date
            const expiresAt = new Date(Date.now() + newTokens.expiresIn * 1000);

            // Save back to repository, including refresh token lifetime if provided
            await this.tokenRepository.saveToken(accountId, {
                accessToken:           newTokens.accessToken,
                refreshToken:          newTokens.refreshToken,
                expiresAt:             expiresAt,
                refreshTokenExpiresIn: newTokens.refreshTokenExpiresIn || null
            });

            console.log('token was refreshed');

            return newTokens.accessToken;

        } catch (error) {
            // If API reports refresh token is invalid or revoked, update DB status
            if (this.isRevocationError(error)) {
                await this.tokenRepository.markDisconnected(accountId);
                throw new OAuthTokenRevokedError(
                    `OAuth connection revoked or expired for ${this.providerName} (${accountId}).`,
                    this.providerName
                );
            }

            // Wrap generic network/temporary failures
            throw new OAuthTokenRefreshError(
                `Failed to refresh ${this.providerName} token: ${error.message}`,
                this.providerName,
                error
            );
        } finally {
            // Ensure lock is released under all circumstances
            releaseLock();
        }
    }

    /**
     * Check if a token expires in less than 5 minutes.
     * @param {Date|null} expiresAt
     * @returns {boolean}
     */
    isExpiringSoon(expiresAt) {
        if (!expiresAt) return true;
        const FIVE_MINUTES_MS = 5 * 60 * 1000;
        return (new Date(expiresAt).getTime() - Date.now()) < FIVE_MINUTES_MS;
    }

    /**
     * Check if a value is in the past (i.e. already expired).
     * Used for the refresh token expiry pre-flight check.
     * @param {Date|null} expiresAt
     * @returns {boolean}
     */
    isAlreadyExpired(expiresAt) {
        if (!expiresAt) return false; // unknown — attempt the refresh anyway
        return new Date(expiresAt).getTime() <= Date.now();
    }

    /**
     * Determine if an error from the OAuth provider signifies token revocation
     * (refresh token invalid, expired, or revoked) rather than a transient
     * failure. On a true positive, the caller stops retrying immediately —
     * no further refresh attempts are made for this connection.
     * @param {Error} error
     * @returns {boolean}
     */
    isRevocationError(error) {
        if (error.response && error.response.data) {
            const data = error.response.data;
            const errStr = JSON.stringify(data).toLowerCase();
            // Typical OAuth revoked response keywords/errors:
            // invalid_grant, invalid_token, unauthorized.
            return errStr.includes('invalid_grant')
                || errStr.includes('invalid_token')
                || errStr.includes('unauthorized');
        }
        return false;
    }
}

module.exports = TokenManager;
