'use strict';

const { Op } = require('sequelize');
const { Notification } = require('../../core/database');

// Notifications are self-expiring: nothing older than this is ever
// returned (or kept) for a user — see _purgeExpired below. Not a fixed
// background job/cron (this app has none registered) — instead this
// purges opportunistically, scoped to the calling user only, every time
// their history is read or added to, which is also the only time an
// expired row could otherwise have been surfaced anyway.
const NOTIFICATION_TTL_MS = 24 * 60 * 60 * 1000; // 24 hours

/**
 * NotificationController
 * -----------------------------------------------------------------
 * Per-user notification history — replaces the old client-only
 * localStorage `fa_notifications` store. Every query below is scoped to
 * req.user.userId (set by the `authenticate` middleware from the
 * verified JWT — see modules/auth/auth.middleware.js) and NEVER a
 * client-suppliable id, so one signed-in user can never read, create
 * under, mark-read, or clear another user's notifications.
 * -----------------------------------------------------------------
 */
class NotificationController {

    /** Hard-deletes this user's notifications older than NOTIFICATION_TTL_MS. */
    async _purgeExpired(userId) {
        await Notification.destroy({
            where: {
                userId,
                createdAt: { [Op.lt]: new Date(Date.now() - NOTIFICATION_TTL_MS) }
            }
        });
    }

    /**
     * GET /api/notifications
     * Returns every (non-expired) notification belonging to the logged-in
     * user, newest first.
     */
    list = async (req, res, next) => {
        try {
            await this._purgeExpired(req.user.userId);

            const notifications = await Notification.findAll({
                where: { userId: req.user.userId },
                order: [['createdAt', 'DESC']]
            });
            res.json({ success: true, notifications });
        } catch (err) {
            next(err);
        }
    };

    /**
     * POST /api/notifications
     * Body: { type, message, detail?, provider? } — validated by
     * schemas.createNotification (core/validation/schemas.js) before
     * this handler ever runs.
     * Creates a notification row owned by the logged-in user.
     */
    create = async (req, res, next) => {
        try {
            const { type, message, detail, provider } = req.body;

            // Opportunistic cleanup — keeps this user's row count bounded
            // even if they never revisit the bell to trigger list()'s purge.
            this._purgeExpired(req.user.userId).catch(() => {});

            const notification = await Notification.create({
                userId: req.user.userId,
                type,
                message,
                detail: detail || null,
                provider: provider || null
            });

            res.status(201).json({ success: true, notification });
        } catch (err) {
            next(err);
        }
    };

    /**
     * PATCH /api/notifications/mark-read
     * Marks the logged-in user's unread notifications as read.
     * Optional body: { ids: [uuid, ...] } to mark only a specific
     * (e.g. currently-visible) subset instead of the whole history.
     */
    markRead = async (req, res, next) => {
        try {
            const where = { userId: req.user.userId, read: false };

            const ids = Array.isArray(req.body?.ids) ? req.body.ids.filter(Boolean) : null;
            if (ids && ids.length) {
                where.id = ids; // Sequelize treats an array value as an IN (...) filter
            }

            await Notification.update({ read: true }, { where });
            res.json({ success: true });
        } catch (err) {
            next(err);
        }
    };

    /**
     * DELETE /api/notifications
     * Clears ("Clear All") every notification belonging to the
     * logged-in user. Returns how many rows were actually removed so the
     * frontend (and server logs) can tell a real 0-row account apart from
     * a delete that silently didn't run.
     */
    clear = async (req, res, next) => {
        try {
            const deletedCount = await Notification.destroy({ where: { userId: req.user.userId } });
            res.json({ success: true, deletedCount });
        } catch (err) {
            next(err);
        }
    };
}

module.exports = new NotificationController();
