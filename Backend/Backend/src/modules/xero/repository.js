const { XeroToken, QuickBooksToken } = require('../../core/database');

/**
 * Every read/write here is scoped to the owning user's email wherever one
 * is supplied. ERP connections must never leak across accounts — a
 * connection belongs exclusively to the email that created it, so callers
 * (the service/controller layers) are expected to always pass the
 * authenticated user's `mail` through to these methods.
 */
class XeroTokenRepository {
    static async getLatestToken(mail) {
        return await XeroToken.findOne({
            where: mail ? { mail } : {},
            order: [['updated_at', 'DESC']]
        });
    }

    /**
     * @param {string} [mail] - Scopes to this user's connections only. Pass
     *   nothing only for genuinely account-agnostic internal use — every
     *   user-facing caller must supply it.
     */
    static async getActiveTokens(mail) {
        return await XeroToken.findAll({
            where: { status: 'Active', ...(mail ? { mail } : {}) },
            order: [['updated_at', 'DESC']]
        });
    }

    /** @param {string} [mail] - Scopes to this user's connections only. */
    static async getAllTokens(mail) {
        return await XeroToken.findAll({
            where: mail ? { mail } : {},
            order: [['created_at', 'DESC']]
        });
    }

    static async upsertToken(tokenData) {
        if (!tokenData.tenant_id) return null;

        return await XeroToken.upsert({
            tenant_id:     tokenData.tenant_id,
            access_token:  tokenData.access_token,
            refresh_token: tokenData.refresh_token,
            expires_in:    tokenData.expires_in,
            token_type:    tokenData.token_type,
            scope:         tokenData.scope,
            session_info:  tokenData.session_info,
            mail:          tokenData.mail,
            company_name:  tokenData.company_name,
            status:        tokenData.status || 'Active'
        });
    }

    /**
     * Deletes only the given user's Xero tokens. `mail` is required —
     * without it this used to `truncate` the entire table (every user's
     * connections), which is exactly the cross-account data loss this
     * scoping exists to prevent, so a missing `mail` is now a no-op rather
     * than a footgun.
     * @param {string} mail
     */
    static async clearTokens(mail) {
        if (!mail) return 0;
        return await XeroToken.destroy({ where: { mail } });
    }
}

module.exports = XeroTokenRepository;
