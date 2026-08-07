const { QuickBooksToken, XeroToken } = require('../../core/database');

/**
 * Every read/write here is scoped to the owning user's email wherever one
 * is supplied. ERP connections must never leak across accounts — a
 * connection belongs exclusively to the email that created it, so callers
 * (the service/controller layers) are expected to always pass the
 * authenticated user's `mail` through to these methods.
 */
class QuickBooksTokenRepository {
    static async getLatestToken(mail) {
        return await QuickBooksToken.findOne({
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
        return await QuickBooksToken.findAll({
            where: { status: 'Active', ...(mail ? { mail } : {}) },
            order: [['updated_at', 'DESC']]
        });
    }

    /** @param {string} [mail] - Scopes to this user's connections only. */
    static async getAllTokens(mail) {
        return await QuickBooksToken.findAll({
            where: mail ? { mail } : {},
            order: [['created_at', 'DESC']]
        });
    }

    static async upsertToken(tokenData) {
        return await QuickBooksToken.upsert(tokenData);
    }

    /**
     * Deletes only the given user's QuickBooks tokens. `mail` is required —
     * without it this used to `truncate` the entire table (every user's
     * connections), which is exactly the cross-account data loss this
     * scoping exists to prevent, so a missing `mail` is now a no-op rather
     * than a footgun.
     * @param {string} mail
     */
    static async clearTokens(mail) {
        if (!mail) return 0;
        return await QuickBooksToken.destroy({ where: { mail } });
    }

    /**
     * Deletes only the given user's Xero tokens. Same `mail`-required
     * safety rule as clearTokens() above.
     * @param {string} mail
     */
    static async clearXeroTokens(mail) {
        if (!mail) return 0;
        return await XeroToken.destroy({ where: { mail } });
    }
}

module.exports = QuickBooksTokenRepository;
