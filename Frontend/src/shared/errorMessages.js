/**
 * Centralized error codes + user-facing copy.
 * -----------------------------------------------------------------
 * Codes here MUST match the backend's standardized error envelope
 * (Backend/Backend/src/core/errors/AppError.js):
 *   { success: false, code, message, details }
 *
 * The frontend never renders `details` (technical/log-only) or a raw
 * `error.message` from an unexpected exception — it always looks up a
 * safe, user-friendly string from this file by `code`, falling back to
 * DEFAULT for anything unrecognized. Add a new entry here whenever the
 * backend introduces a new error code.
 * -----------------------------------------------------------------
 */

export const ERROR_CODES = {
    CONNECTION_REFUSED:  "ERR_CONNECTION_REFUSED",
    SESSION_EXPIRED:      "ERR_SESSION_EXPIRED",
    ERP_SESSION_EXPIRED:  "ERR_ERP_SESSION_EXPIRED",
    ERP_REFRESH_FAILED:   "ERR_ERP_REFRESH_FAILED",
    QB_SUBSCRIPTION_EXPIRED: "ERR_QB_SUBSCRIPTION_EXPIRED",
    VALIDATION:           "ERR_VALIDATION",
    UNAUTHORIZED:         "ERR_UNAUTHORIZED",
    FORBIDDEN:            "ERR_FORBIDDEN",
    NOT_FOUND:            "ERR_NOT_FOUND",
    CONFLICT:             "ERR_CONFLICT",
    LIMIT_REACHED:        "ERR_LIMIT_REACHED",
    INTERNAL:             "ERR_INTERNAL",
    UNKNOWN:               "ERR_UNKNOWN"
};

export const ERROR_MESSAGES = {
    [ERROR_CODES.CONNECTION_REFUSED]:
        "Offline: Cannot connect to the server. Please check your internet connection.",
    [ERROR_CODES.SESSION_EXPIRED]:
        "Session expired. Please log in again.",
    [ERROR_CODES.ERP_SESSION_EXPIRED]:
        "Your QuickBooks/Xero session has expired. Please reconnect.",
    [ERROR_CODES.ERP_REFRESH_FAILED]:
        "Could not reach QuickBooks/Xero right now. Please try again shortly.",
    [ERROR_CODES.QB_SUBSCRIPTION_EXPIRED]:
        "Your QuickBooks subscription has expired or been suspended. Please log into QuickBooks to update your billing.",
    [ERROR_CODES.VALIDATION]:
        "That request looks incomplete. Please check the form and try again.",
    [ERROR_CODES.UNAUTHORIZED]:
        "Invalid credentials. Please try again.",
    [ERROR_CODES.FORBIDDEN]:
        "You don't have permission to do that.",
    [ERROR_CODES.NOT_FOUND]:
        "We couldn't find what you were looking for.",
    [ERROR_CODES.CONFLICT]:
        "That conflicts with existing data.",
    [ERROR_CODES.LIMIT_REACHED]:
        "You've reached your plan's limit.",
    [ERROR_CODES.INTERNAL]:
        "Something went wrong on our end. Please try again later.",
    DEFAULT:
        "Something went wrong. Please try again."
};

/** Looks up a safe, user-facing message for a backend error code. */
export function getFriendlyMessage(code) {
    return ERROR_MESSAGES[code] || ERROR_MESSAGES.DEFAULT;
}
