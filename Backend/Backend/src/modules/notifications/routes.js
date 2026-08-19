'use strict';

const express    = require('express');
const router     = express.Router();
const controller = require('./controller');
const { authenticate } = require('../auth/auth.middleware');
const { validate } = require('../../core/middleware/validate');
const schemas = require('../../core/validation/schemas');

/**
 * Notifications Routes
 * -----------------------------------------------------------------
 * Mounted at /api/notifications (see src/routes/index.js). Every route
 * is JWT-protected — `authenticate` populates req.user.userId from the
 * verified token, which is what scopes every read/write to the calling
 * user only (see NotificationController).
 * -----------------------------------------------------------------
 */

// GET /api/notifications — fetch all notifications for the logged-in user
router.get('/', authenticate, controller.list);

// POST /api/notifications — create a notification for the logged-in user
router.post('/', authenticate, validate(schemas.createNotification), controller.create);

// PATCH /api/notifications/mark-read — mark all (or visible) notifications as read
router.patch('/mark-read', authenticate, controller.markRead);

// DELETE /api/notifications — clear all notifications for the logged-in user
router.delete('/', authenticate, controller.clear);

module.exports = router;
