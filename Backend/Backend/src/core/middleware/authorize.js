'use strict';

const { AuthorizationError } = require('../errors/AppError');

/**
 * authorize(...allowedRoles)
 * ----------------------------------------------------------------
 * Authorization Validation (role-based access control).
 *
 * FinAccrual previously had NO role check anywhere — `role` was carried
 * in the JWT payload (both the main user token from
 * modules/auth/jwt.service.js and the admin token from
 * modules/admin/controller.js) and stored on the User/Admin models, but
 * nothing ever read `req.user.role` to gate a route. That's an
 * authentication/authorization gap: `authenticate` (modules/auth/
 * auth.middleware.js) and `authenticateJWT` (core/middleware/
 * authMiddleware.js) only prove the caller has SOME valid FinAccrual
 * JWT — they don't check that JWT is allowed to do THIS operation.
 *
 * Concretely, this was exploitable: both the main-app JWT and the admin
 * JWT are signed with the same JWT_SECRET, so a regular user's own JWT
 * (role: 'user') passed signature verification against the admin
 * module's `authenticateJWT` middleware just fine — GET /api/admin/me
 * had no role check behind that middleware, so a signed-in FinAccrual
 * user could hit an admin-module endpoint with their own token. Adding
 * `authorize('admin')` there closes that gap.
 *
 * Must run AFTER an authentication middleware (`authenticate` or
 * `authenticateJWT`) that has already set `req.user`.
 *
 * Usage:
 *   router.get('/admin/users', authenticateJWT, authorize('admin'), controller.listUsers);
 * ----------------------------------------------------------------
 * @param {...string} allowedRoles - one or more roles permitted through
 */
function authorize(...allowedRoles) {
    return (req, res, next) => {
        const role = req.user && req.user.role;

        if (!role || !allowedRoles.includes(role)) {
            return next(new AuthorizationError(allowedRoles.join(' or ')));
        }

        next();
    };
}

module.exports = { authorize };
