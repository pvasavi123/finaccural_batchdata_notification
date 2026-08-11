const express = require("express");

const router = express.Router();

const AdminController = require('./controller');
const authenticateJWT = require('../../core/middleware/authMiddleware');
const { authorize } = require('../../core/middleware/authorize');
const { validate } = require('../../core/middleware/validate');
const schemas = require('../../core/validation/schemas');
const { authLimiter } = require('../../core/middleware/rateLimiters');

// Request Schema Validation + Security Validation: same Joi-schema and
// brute-force-limiter treatment as the main /api/auth/login|signup.
router.post("/login", authLimiter, validate(schemas.adminLogin), AdminController.login);
router.post("/signup", authLimiter, validate(schemas.adminSignup), AdminController.signup);

// Authorization Validation: authenticateJWT only proves the caller has
// SOME valid FinAccrual JWT (main-app and admin tokens share the same
// JWT_SECRET) — authorize('admin') is what actually restricts this to
// admins. Previously missing, so any signed-in FinAccrual user's own
// token would have passed straight through to this route.
router.get("/me", authenticateJWT, authorize('admin'), (req, res) => {
    res.json({
        success: true,
        admin: req.user
    });
});

// GET /api/admin/users — admin-only list of every FinAccrual user
// account. See AdminController.listUsers for details.
router.get("/users", authenticateJWT, authorize('admin'), AdminController.listUsers);

module.exports = router;
