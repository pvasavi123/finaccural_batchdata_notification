/**
 * FinAccrual Excel Add-in Taskpane Application
 *
 * Architecture: Modern Multi-View SaaS Design
 *
 * Namespaces:
 *  - AppState      : Centralized reactive application state
 *  - ViewRouter    : Multi-view screen navigation controller
 *  - ApiService    : Backend API calls (auth, subscription, ERP tokens)
 *  - ExcelService  : Excel JS workbook sheet management
 *  - AuthService   : Google / Microsoft OAuth popup handlers
 *  - DashboardService : Dashboard UI rendering and ERP operations
 *  - NotificationService : Bell icon / drawer notification history
 *  - AppController : Event binding, session restoration, init
 */

import { writeRowsInBatches } from "./batchDataLoader.js";
import {
    ERROR_CODES,
    ApiError,
    parseApiError,
    networkError,
    showBanner,
    hideBanner,
    showToast
} from "../shared/apiErrorHandler.js";
import { getFriendlyMessage } from "../shared/errorMessages.js";

Office.onReady(() => {

    // ============================================================
    // 1. CENTRALIZED APPLICATION STATE
    // ============================================================
    const AppState = {
        // Auth / Subscription
        userEmail:        localStorage.getItem("fa_user_email")  || null,
        userName:         localStorage.getItem("fa_user_name")   || null,
        userProvider:     localStorage.getItem("fa_user_provider") || null,
        hasSubscription:  localStorage.getItem("fa_has_subscription") === "true",
        subscriptionId:   localStorage.getItem("fa_subscription_id")   || null,
        subscriptionPlan: (v => (!v || v === 'null' || v === 'undefined') ? null : v)(localStorage.getItem("fa_subscription_plan")),

        // Pending checkout details (set when user selects a plan)
        pendingPlan:      null,
        pendingPrice:     null,
        pendingCycle:     null,

        // ERP Connection
        erpConnected:     localStorage.getItem("fa_erp_connected") === "true",
        erpType:          localStorage.getItem("fa_erp_type") || null,        // "quickbooks" | "xero"
        erpOrgName:       localStorage.getItem("fa_erp_org")  || null,
        erpConnectedDate: localStorage.getItem("fa_erp_date") || null,
        forceWelcome:     false,

        // JWT token — persisted across sessions
        jwtToken:         localStorage.getItem("fa_jwt_token") || null,

        // Set true the moment the backend reports ERR_SESSION_EXPIRED.
        // While true, ApiService.apiFetch refuses to make further requests
        // (avoids hammering the backend with a storm of repeated 401s)
        // until a fresh token is obtained via login.
        sessionExpired:   false,

        // ERP Operations
        currentProvider: "quickbooks",
        get currentTier() {
            const plan = (AppState.subscriptionPlan || "pro").toLowerCase();
            if (plan.includes("basic")) return "basic";
            if (plan.includes("standard")) return "standard";
            return "pro";
        },
        connectionId:    null,
        isConnected:     false
    };

    // ============================================================
    // 2. VIEW ROUTER
    // ============================================================
    const ViewRouter = {
        /**
         * Shows a view by its ID name (e.g. "Welcome", "Dashboard")
         * @param {string} name - View name that maps to #view<Name>
         */
        show(name) {
            document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
            const el = document.getElementById("view" + name);
            if (el) el.classList.add("active");
            if (name && name !== "Loading" && name !== "Error") {
                localStorage.setItem("fa_last_view", name);
            }
            if (name === "Plans") {
                const nameVal = AppState.userName || AppState.userEmail || "User";
                const initial = nameVal.charAt(0).toUpperCase();
                const avatarEl = document.getElementById("plansUserAvatar");
                const emailEl = document.getElementById("plansUserEmail");
                if (avatarEl) avatarEl.textContent = initial;
                if (emailEl) emailEl.textContent = AppState.userEmail || "";
            }
        }
    };

    // ============================================================
    // NOTIFICATION SERVICE — Bell icon + drawer notification history
    // ============================================================
    const NotificationService = {
        STORAGE_KEY: "fa_notifications",
        MAX_ITEMS: 50,
        _lastNotif: null,

        /**
         * Normalizes a provider tag to "quickbooks" | "xero" | null.
         * null means "global" — not tied to either ERP, so it's always
         * visible regardless of which provider the user is currently on
         * (e.g. login, payment, logout).
         * @param {string} [provider]
         * @returns {"quickbooks"|"xero"|null}
         */
        _normalizeProvider(provider) {
            if (provider === "quickbooks" || provider === "xero") return provider;
            return null;
        },

        /**
         * Filters a notification list down to what's visible in the
         * current provider context: global (provider-less) entries plus
         * whichever ERP the user is actively using. This is the single
         * choke point that keeps QuickBooks and Xero notifications
         * (toast, badge count, and drawer history) fully separated —
         * QuickBooks notifications never surface while on Xero and vice
         * versa.
         * @param {Array} list
         * @returns {Array}
         */
        _forCurrentContext(list) {
            const ctx = (typeof AppState !== "undefined" && AppState.currentProvider) || null;
            return list.filter(n => !n.provider || n.provider === ctx);
        },

        /** @returns {Array<{id:string,type:'success'|'error',message:string,detail:string,provider:('quickbooks'|'xero'|null),timestamp:string,read:boolean}>} */
        _load() {
            try {
                const raw = localStorage.getItem(this.STORAGE_KEY);
                const parsed = raw ? JSON.parse(raw) : [];
                return Array.isArray(parsed) ? parsed : [];
            } catch (_) {
                return [];
            }
        },

        _save(list) {
            try {
                localStorage.setItem(this.STORAGE_KEY, JSON.stringify(list));
            } catch (_) {
                // Storage full/unavailable — notifications just won't persist
                // across a refresh, but the app keeps working.
            }
        },

        _escapeHtml(str) {
            const div = document.createElement("div");
            div.textContent = String(str == null ? "" : str);
            return div.innerHTML;
        },

        _formatTimestamp(iso) {
            const d = new Date(iso);
            if (isNaN(d.getTime())) return "";
            const now = new Date();
            const timeStr = d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
            if (d.toDateString() === now.toDateString()) return timeStr;
            const dateStr = d.toLocaleDateString([], { month: "short", day: "numeric" });
            return `${dateStr}, ${timeStr}`;
        },

        /**
         * Records a new notification, shows it immediately as a top-right
         * toast (the primary way the user learns an action succeeded or
         * failed — the bell is just the persisted history, never required
         * reading), and refreshes the bell badge. If the drawer happens to
         * already be open, the new entry is shown there too and marked read.
         *
         * QuickBooks/Xero separation: `provider` tags which ERP the action
         * belongs to. A tagged notification only ever toasts, counts toward
         * the badge, or appears in the drawer while the user is actively on
         * that same provider (AppState.currentProvider) — it's still saved
         * to history so it surfaces correctly if the user switches back
         * later, but it's fully invisible in the meantime. Untagged (global)
         * notifications — login, payment, logout, etc. — aren't tied to
         * either ERP and always show.
         * @param {string} message - short outcome, e.g. "Data pull completed successfully"
         * @param {"success"|"error"} type
         * @param {string} [detail] - optional second line, e.g. "Data synchronized successfully."
         * @param {"quickbooks"|"xero"} [provider] - which ERP this belongs to, if any
         */
        add(message, type, detail, provider) {
            if (!message) return;
            const normalizedType = type === "error" ? "error" : "success";
            const normalizedProvider = this._normalizeProvider(provider);

            // Duplicate guard — the same completed outcome can occasionally
            // be reported twice (e.g. two callback paths both observing the
            // same finished operation). Suppress an identical repeat within
            // a short window so each completed action produces exactly one
            // toast/notification, never more.
            const now = Date.now();
            if (
                this._lastNotif &&
                this._lastNotif.message === message &&
                this._lastNotif.type === normalizedType &&
                this._lastNotif.provider === normalizedProvider &&
                (now - this._lastNotif.at) < 1500
            ) {
                return;
            }
            this._lastNotif = { message, type: normalizedType, provider: normalizedProvider, at: now };

            const list = this._load();
            list.unshift({
                id: `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
                type: normalizedType,
                message: String(message),
                detail: detail ? String(detail) : "",
                provider: normalizedProvider,
                timestamp: new Date().toISOString(),
                read: false
            });
            if (list.length > this.MAX_ITEMS) list.length = this.MAX_ITEMS;
            this._save(list);

            // Only surface it (toast + badge/drawer refresh) if it's global
            // or matches the ERP the user is currently on — a QuickBooks
            // notification must never appear while on Xero, and vice versa.
            const currentCtx = (typeof AppState !== "undefined" && AppState.currentProvider) || null;
            const isVisibleNow = !normalizedProvider || normalizedProvider === currentCtx;
            if (!isVisibleNow) {
                return;
            }

            // Immediate feedback — always, regardless of whether the drawer
            // is open or the bell has ever been clicked.
            this.showToast(message, normalizedType, detail);

            const drawer = document.getElementById("notifDrawer");
            if (drawer && drawer.style.display !== "none") {
                // Drawer is open — the new notification is on screen, so
                // treat it as read immediately instead of leaving a stray
                // unread badge behind an already-open drawer.
                this.markAllRead();
                this.renderDrawer();
            } else {
                this.renderBadge();
            }
        },

        /**
         * Renders a single floating, top-right toast card. Multiple toasts
         * stack (most recent on top); each auto-dismisses on its own timer
         * but can also be closed early via the × button.
         * @param {string} message
         * @param {"success"|"error"} type
         * @param {string} [detail]
         * @param {number} [duration=4500]
         */
        showToast(message, type, detail, duration = 4500) {
            const container = document.getElementById("notifToastContainer");
            if (!container) return;

            const toastEl = document.createElement("div");
            toastEl.className = `fa-notif-toast fa-notif-toast-${type === "error" ? "error" : "success"}`;
            toastEl.innerHTML = `
                <span class="fa-notif-toast-icon">${type === "error" ? "✕" : "✓"}</span>
                <div class="fa-notif-toast-body">
                    <div class="fa-notif-toast-title">${this._escapeHtml(message)}</div>
                    ${detail ? `<div class="fa-notif-toast-detail">${this._escapeHtml(detail)}</div>` : ""}
                </div>
                <button class="fa-notif-toast-close" aria-label="Close">&times;</button>
            `;
            container.appendChild(toastEl);

            let dismissTimer = null;
            const removeToast = () => {
                clearTimeout(dismissTimer);
                toastEl.classList.add("fa-notif-toast-hide");
                setTimeout(() => toastEl.remove(), 200);
            };

            toastEl.querySelector(".fa-notif-toast-close")?.addEventListener("click", removeToast);
            dismissTimer = setTimeout(removeToast, duration);
        },

        /**
         * Re-renders the badge (and the drawer, if it's open) against the
         * current provider context. Call this whenever AppState.currentProvider
         * changes (switching between QuickBooks and Xero) so a badge count
         * or open drawer left over from the other provider doesn't linger —
         * it recomputes to show only what's actually visible now.
         */
        refreshForContext() {
            this.renderBadge();
            const drawer = document.getElementById("notifDrawer");
            if (drawer && drawer.style.display !== "none") {
                this.renderDrawer();
            }
        },

        /** Notifications visible right now — global entries plus the active provider's. */
        getAll() {
            return this._forCurrentContext(this._load());
        },

        getUnreadCount() {
            return this._forCurrentContext(this._load()).filter(n => !n.read).length;
        },

        /**
         * Marks read only the notifications currently visible (global +
         * active provider) — a QuickBooks notification the user hasn't
         * seen yet (because they're on Xero) stays unread until they
         * actually switch to QuickBooks and see it.
         */
        markAllRead() {
            const list = this._load();
            const ctx = (typeof AppState !== "undefined" && AppState.currentProvider) || null;
            let changed = false;
            list.forEach(n => {
                if (!n.read && (!n.provider || n.provider === ctx)) {
                    n.read = true;
                    changed = true;
                }
            });
            if (changed) this._save(list);
            this.renderBadge();
        },

        /** Permanently deletes every notification from storage. */
        clearAll() {
            localStorage.removeItem(this.STORAGE_KEY);
            this.renderBadge();
            this.renderDrawer();
        },

        renderBadge() {
            const badge = document.getElementById("notifBadge");
            if (!badge) return;
            const count = this.getUnreadCount();
            if (count > 0) {
                badge.textContent = count > 99 ? "99+" : String(count);
                badge.style.display = "flex";
            } else {
                badge.style.display = "none";
            }
        },

        renderDrawer() {
            const listEl = document.getElementById("notifList");
            if (!listEl) return;
            const notifications = this.getAll();

            if (notifications.length === 0) {
                listEl.innerHTML = `<div class="fa-notif-empty">No notifications available.</div>`;
                return;
            }

            listEl.innerHTML = notifications.map(n => {
                const icon = n.type === "error" ? "✕" : "✓";
                return `
                    <div class="fa-notif-item fa-notif-item-${n.type === "error" ? "error" : "success"}">
                        <span class="fa-notif-icon">${icon}</span>
                        <div class="fa-notif-content">
                            <div class="fa-notif-message">${this._escapeHtml(n.message)}</div>
                            ${n.detail ? `<div class="fa-notif-detail">${this._escapeHtml(n.detail)}</div>` : ""}
                            <div class="fa-notif-time">${this._formatTimestamp(n.timestamp)}</div>
                        </div>
                    </div>
                `;
            }).join("");
        },

        /** Wires up the bell button, drawer, and Clear All. Call once on init. */
        init() {
            this.renderBadge();

            const drawer = () => document.getElementById("notifDrawer");

            const toggleDrawer = (e) => {
                e.stopPropagation();
                const el = drawer();
                if (!el) return;
                const willOpen = el.style.display === "none";
                el.style.display = willOpen ? "flex" : "none";
                if (willOpen) {
                    // Opening the drawer shows the full history and marks
                    // everything read — the unread badge disappears.
                    this.renderDrawer();
                    this.markAllRead();
                }
            };
            document.getElementById("notifBellBtn")?.addEventListener("click", toggleDrawer);

            // Close the drawer when clicking anywhere outside it (or the bell).
            document.addEventListener("click", (e) => {
                const el = drawer();
                const bell = document.getElementById("notifBellBtn");
                if (el && el.style.display !== "none" && !el.contains(e.target) && !(bell && bell.contains(e.target))) {
                    el.style.display = "none";
                }
            });

            // Clear All — deletes immediately, no confirmation dialog, then
            // closes the drawer. Stop propagation so this click doesn't
            // also trigger the "click outside closes the drawer" listener
            // above (that's a no-op here anyway since we close it
            // ourselves, but keeps behavior explicit/predictable).
            document.getElementById("notifClearAllBtn")?.addEventListener("click", (e) => {
                e.stopPropagation();
                this.clearAll();
                const el = drawer();
                if (el) el.style.display = "none";
            });
        }
    };

    // ============================================================
    // 3. API SERVICE LAYER
    // ============================================================
    const ApiService = {
        // ── Single source-of-truth for the backend base URL ────────────
        BASE: "http://localhost:8080",

        /**
         * Centralized error reaction — every call through apiFetch funnels
         * here. Branches on the standardized backend `code` field (never on
         * message text) and reacts globally:
         *   - ERR_CONNECTION_REFUSED  -> red offline banner + Retry
         *   - ERR_SESSION_EXPIRED     -> toast + clear tokens + redirect to
         *                                 Login + block further requests
         *   - ERR_ERP_SESSION_EXPIRED -> orange banner + Reconnect
         * Any other code is left to the caller's own try/catch — this only
         * reacts to the three centrally-handled scenarios.
         * @param {ApiError} apiErr
         * @param {{ retry?: () => void }} [opts]
         */
        handleGlobalApiError(apiErr, opts = {}) {
            switch (apiErr.code) {
                case ERROR_CODES.CONNECTION_REFUSED: {
                    showBanner({
                        type: "offline",
                        message: apiErr.message,
                        actionLabel: "Retry",
                        onAction: () => {
                            hideBanner();
                            if (opts.retry) opts.retry();
                        }
                    });
                    break;
                }
                case ERROR_CODES.SESSION_EXPIRED: {
                    // Avoid re-triggering the redirect/toast for every
                    // in-flight request that fails after the first 401.
                    if (AppState.sessionExpired) break;
                    AppState.sessionExpired = true;
                    showToast(getFriendlyMessage(ERROR_CODES.SESSION_EXPIRED));
                    AuthService.logout();
                    break;
                }
                case ERROR_CODES.ERP_SESSION_EXPIRED: {
                    const provider = AppState.currentProvider === "xero" ? "Xero" : "QuickBooks";
                    showBanner({
                        type: "erp",
                        message: apiErr.message,
                        actionLabel: "Reconnect",
                        onAction: () => {
                            hideBanner();
                            DashboardService.launchERPOAuth(AppState.currentProvider === "xero" ? "xero" : "quickbooks");
                        }
                    });
                    void provider; // reserved for future provider-specific copy
                    break;
                }
                default:
                    // Not one of the three centrally-handled scenarios —
                    // caller's own catch block is responsible for the UI.
                    break;
            }
        },

        /**
         * Authenticated fetch helper — the single choke point every API call
         * in this app goes through. Automatically attaches the JWT Bearer
         * token, and centrally reacts to the app's three standardized error
         * scenarios (offline/unreachable backend, expired session, expired
         * ERP connection) via handleGlobalApiError above, in addition to
         * returning the raw Response so existing callers keep working
         * unchanged (their own res.ok / res.json() handling still applies).
         *
         * Usage (same API as window.fetch):
         *   const res = await ApiService.apiFetch('/api/connections', { method: 'GET' });
         *
         * @param {string}  path     - Relative (/api/...) or absolute URL
         * @param {object}  options  - Standard fetch options (method, body, headers, …)
         * @returns {Promise<Response>}
         */
        async apiFetch(path, options = {}) {
            // Once a session-expired redirect has fired, refuse further
            // requests until a fresh token is obtained via login — prevents
            // a storm of repeated 401s while the user is on the Login view.
            if (AppState.sessionExpired) {
                throw new ApiError(ERROR_CODES.SESSION_EXPIRED, getFriendlyMessage(ERROR_CODES.SESSION_EXPIRED), "Blocked: session already marked expired.");
            }

            const url = path.startsWith("http") ? path : `${this.BASE}${path}`;
            const headers = { ...(options.headers || {}) };
            if (AppState.jwtToken) {
                headers["Authorization"] = `Bearer ${AppState.jwtToken}`;
            }

            const retry = () => this.apiFetch(path, options);

            let res;
            try {
                res = await fetch(url, { ...options, headers });
            } catch (networkErr) {
                const apiErr = networkError(networkErr);
                this.handleGlobalApiError(apiErr, { retry });
                throw apiErr;
            }

            if (!res.ok) {
                const apiErr = await parseApiError(res.clone());
                this.handleGlobalApiError(apiErr, { retry });
                // Rethrow-free: callers keep doing their own `res.ok` /
                // `res.json()` handling exactly as before. The global
                // reaction above already fired as a side effect.
                return res;
            }

            // A successful call clears any stale offline banner.
            hideBanner();
            return res;
        },

        /**
         * Checks subscription status for the given email from backend.
         * Also stores the JWT token if the server returns one.
         *
         * If the backend is unreachable (no internet / server down), this
         * surfaces the same offline banner + Retry action as every other
         * call in the app instead of failing silently — the caller still
         * gets a graceful fallback value so existing view-routing logic
         * keeps working, but the user now sees *why*.
         */
        async checkSubscription(email) {
            let res;
            try {
                res = await fetch(`${this.BASE}/api/auth/login`, {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ email })
                });
            } catch (networkErr) {
                const apiErr = networkError(networkErr);
                this.handleGlobalApiError(apiErr, { retry: () => this.checkSubscription(email) });
                return { hasSubscription: AppState.hasSubscription };
            }

            if (!res.ok) {
                const apiErr = await parseApiError(res.clone());
                this.handleGlobalApiError(apiErr);
                return { hasSubscription: false };
            }

            const result = await res.json();
            // Persist the token if the backend returned one
            if (result.token) {
                AppState.jwtToken = result.token;
                AppState.sessionExpired = false; // fresh token — lift the request block
                localStorage.setItem("fa_jwt_token", result.token);
            }
            hideBanner();
            return result;
        },

        /** Checks ERP token validity from backend (JWT required). */
        async checkTokens(provider) {
            const path = provider === "quickbooks"
                ? "/api/quickbooks/tokens/"
                : "/api/xero/tokens";
            const res = await this.apiFetch(path);
            if (!res.ok) throw new Error(`Failed to check ${provider} tokens.`);
            return await res.json();
        },

        /** Disconnects the ERP provider from backend (JWT required). */
        async disconnectERP(provider) {
            const path = provider === "quickbooks"
                ? "/api/quickbooks/disconnect"
                : "/api/xero/disconnect";
            const res = await this.apiFetch(path, { method: "POST" });
            if (!res.ok) throw new Error(`Failed to disconnect ${provider}.`);
            return await res.json();
        },

        /**
         * Pulls master metadata from ERP APIs via the unified backend pull endpoint.
         * @param {string} provider  - The active ERP provider ("quickbooks" | "xero")
         * @param {string} companyId - Selected company identifier
         * @returns {Promise<object>} Map of company, customers, vendors, accounts, classes, locations
         */
        async fetchMasterData(provider, companyId) {
            const params = new URLSearchParams({
                companyId: companyId || "",
                platform:  provider || "",
                tier:      AppState.currentTier || ""
            });
            const res = await this.apiFetch(`/api/pull-master-data?${params.toString()}`, {
                method: "GET"
            });
            if (!res.ok) {
                // apiFetch already triggered the global reaction (offline
                // banner / ERP-expired banner) as a side effect above.
                // Re-parse here too (parseApiError clones internally, so
                // this doesn't double-consume the body) so this call's own
                // catch block can branch on `.code` instead of guessing
                // from message text.
                throw await parseApiError(res);
            }
            const data = await res.json();

            if (data && data.tokenRefreshed) {
                DashboardService.addLog("Token is refreshed");
            }

            // This request/response happens entirely in the browser (fetch
            // API running in the taskpane WebView), so the count is logged
            // to the browser console, not a terminal.
            const counts = {
                company: Array.isArray(data.company) ? data.company.length : (data.company ? 1 : 0),
                accounts: Array.isArray(data.accounts) ? data.accounts.length : 0,
                classes: Array.isArray(data.classes) ? data.classes.length : 0,
                locations: Array.isArray(data.locations) ? data.locations.length : 0,
                customers: Array.isArray(data.customers) ? data.customers.length : 0,
                vendors: Array.isArray(data.vendors) ? data.vendors.length : 0
            };
            const total = counts.company + counts.accounts + counts.classes + counts.locations + counts.customers + counts.vendors;
            console.log(
                `[FinAccrual] Master data records for company ${companyId} (${provider}): ` +
                `accounts=${counts.accounts}, classes=${counts.classes}, locations=${counts.locations}, ` +
                `customers=${counts.customers}, vendors=${counts.vendors}, company=${counts.company}, total=${total}`
            );

            return data;
        }
    };

    // ============================================================
    // 4. EXCEL SERVICE LAYER
    // ============================================================

    /**
     * A record counts as "changed since the last sync" if it's either
     * newly added (isNew) or a pre-existing record that was modified
     * (isUpdated). Both are treated identically for display purposes:
     * moved to the bottom of the sheet and highlighted.
     * @param {{isNew?: boolean, isUpdated?: boolean}} item
     * @returns {boolean}
     */
    function isChangedRecord(item) {
        return !!(item.isNew || item.isUpdated);
    }

    /**
     * Reorders a list so unchanged existing records keep their original
     * relative order and are written first (default background, same
     * position), with new-or-updated records appended after them —
     * also preserving their own relative order. Array.prototype.filter
     * preserves order, so changed records are never interleaved among
     * the unchanged ones.
     *
     * When nothing has changed, this is a no-op: the list comes back in
     * its original order and every row keeps the default background.
     *
     * @param {Array<{isNew?: boolean, isUpdated?: boolean}>} list
     * @returns {Array}
     */
    function partitionExistingThenNew(list) {
        const unchanged = list.filter(item => !isChangedRecord(item));
        const changed = list.filter(item => isChangedRecord(item));
        return unchanged.concat(changed);
    }

    /**
     * Counts new-or-updated records across an entire Master Data Pull
     * response, used to enrich the "pull succeeded" status/notification
     * message with e.g. "3 new/updated records found".
     * @param {object} data - Master data payload (accounts/classes/locations/customers/vendors)
     * @returns {number}
     */
    function countChangedMasterDataRecords(data) {
        if (!data) return 0;
        const sections = [data.accounts, data.classes, data.locations, data.customers, data.vendors];
        return sections.reduce((total, list) => {
            return Array.isArray(list) ? total + list.filter(isChangedRecord).length : total;
        }, 0);
    }

    const ExcelService = {
        /**
         * Populates Excel workbook sheet "1.Master_Data" with ERP payload.
         * @param {string} provider - Active ERP provider
         * @param {object} data - Master data payload
         */
        async writeMasterData(provider, data) {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("1.Master_Data");

                // Clear existing records in spreadsheet grid below headers.
                // "All" (not just "Contents") is required so any "new record"
                // highlight fill from a previous refresh doesn't linger on
                // rows that are no longer new — font size/wrap and column
                // width are reapplied unconditionally below anyway.
                const clearRange = sheet.getRange("A2:AB10000");
                clearRange.clear("All");

                // Group all data by Organization
                const orgGroupsMap = new Map();
                const fallbackOrgName = provider === "quickbooks" ? "QuickBooks Company" : "Xero Organisation";

                const getOrCreateGroup = (id, name) => {
                    const cleanName = (name && name !== "Default" && name !== "Default Organization") ? name : fallbackOrgName;
                    const key       = cleanName.trim();

                    if (!orgGroupsMap.has(key)) {
                        orgGroupsMap.set(key, {
                            id: cleanName,
                            name: cleanName,
                            accounts: [],
                            classes: [],
                            locations: [],
                            entities: []
                        });
                    }
                    return orgGroupsMap.get(key);
                };

                // 1. Group Companies
                const rawCompanies = Array.isArray(data.company)
                    ? data.company
                    : (data.company ? [data.company] : []);

                for (const c of rawCompanies) {
                    if (c) getOrCreateGroup(c.name || c.id, c.name);
                }

                // 2. Group Accounts
                if (data.accounts) {
                    for (const a of data.accounts) {
                        const nameKey = a.clientName || a.clientId || fallbackOrgName;
                        const group = getOrCreateGroup(a.clientId, nameKey);
                        group.accounts.push(a);
                    }
                }

                // 3. Group Classes
                if (data.classes) {
                    for (const c of data.classes) {
                        const nameKey = c.clientName || c.clientId || fallbackOrgName;
                        const group = getOrCreateGroup(c.clientId, nameKey);
                        group.classes.push(c);
                    }
                }

                // 4. Group Locations
                if (data.locations) {
                    for (const l of data.locations) {
                        const nameKey = l.clientName || l.clientId || fallbackOrgName;
                        const group = getOrCreateGroup(l.clientId, nameKey);
                        group.locations.push(l);
                    }
                }

                // 5. Group Entities (Customers and Vendors)
                if (data.customers) {
                    for (const cust of data.customers) {
                        const nameKey = cust.clientName || cust.clientId || fallbackOrgName;
                        const group = getOrCreateGroup(cust.clientId, nameKey);
                        group.entities.push({
                            clientId: nameKey,
                            name: cust.name || cust.DisplayName || cust.Name || "",
                            type: "Customer",
                            id: cust.id || cust.Id || cust.ContactID || "",
                            status: cust.active !== undefined ? (cust.active ? "Active" : "Inactive") : "Active",
                            isNew: !!cust.isNew,
                            isUpdated: !!cust.isUpdated
                        });
                    }
                }

                if (data.vendors) {
                    for (const vend of data.vendors) {
                        const nameKey = vend.clientName || vend.clientId || fallbackOrgName;
                        const group = getOrCreateGroup(vend.clientId, nameKey);
                        group.entities.push({
                            clientId: nameKey,
                            name: vend.name || vend.DisplayName || vend.Name || "",
                            type: "Vendor",
                            id: vend.id || vend.Id || vend.ContactID || "",
                            status: vend.active !== undefined ? (vend.active ? "Active" : "Inactive") : "Active",
                            isNew: !!vend.isNew,
                            isUpdated: !!vend.isUpdated
                        });
                    }
                }

                // Write grouped data sequentially with 2 blank rows between organizations
                let currentRow = 2;

                for (const [key, group] of orgGroupsMap) {
                    const orgName = group.name; // Client ID and Client Name are identical

                    const accCount    = group.accounts.length;
                    const classCount  = group.classes.length;
                    const locCount    = group.locations.length;
                    const entityCount = group.entities.length;

                    const maxRows = Math.max(1, accCount, classCount, locCount, entityCount);

                    // Section 1 (A:B) - Client Config (Client ID and Client Name SAME)
                    sheet.getRange(`A${currentRow}:B${currentRow}`).values = [[orgName, orgName]];

                    // Section 2 (D:L) - Accounts, written in batches of
                    // MASTER_DATA_WRITE_BATCH_SIZE rows so large chart-of-accounts
                    // payloads don't sit in memory as one giant pending write.
                    if (accCount > 0) {
                        // Unchanged accounts keep their original order and
                        // default background; new-or-updated ones are moved
                        // to the bottom and highlighted — never interleaved
                        // in the middle of the existing rows.
                        const orderedAccounts = partitionExistingThenNew(group.accounts);
                        const accValues = orderedAccounts.map(a => [
                            orgName,
                            a.acctNum || a.code || a.AcctNum || a.Code || "",
                            a.name || a.Name || "",
                            a.accountType || a.type || a.AccountType || a.Type || "",
                            a.accountSubType || a.description || a.AccountSubType || "",
                            a.classification || a.Classification || "",
                            a.fullyQualifiedName || a.name || a.Name || "",
                            a.active !== undefined ? (a.active ? "Active" : "Inactive") : "Active",
                            a.id || a.Id || a.AccountID || ""
                        ]);
                        const accHighlightFlags = orderedAccounts.map(isChangedRecord);
                        await writeRowsInBatches(context, sheet, "D", "L", currentRow, accValues, accHighlightFlags);
                    }

                    // Section 3 (N:Q) - Classes, batched the same way.
                    if (classCount > 0) {
                        const orderedClasses = partitionExistingThenNew(group.classes);
                        const classValues = orderedClasses.map(c => [
                            orgName,
                            c.name || c.Name || "",
                            c.id || c.Id || "",
                            c.active !== undefined ? (c.active ? "Active" : "Inactive") : "Active"
                        ]);
                        const classHighlightFlags = orderedClasses.map(isChangedRecord);
                        await writeRowsInBatches(context, sheet, "N", "Q", currentRow, classValues, classHighlightFlags);
                    }

                    // Section 4 (S:V) - Locations, batched the same way.
                    if (locCount > 0) {
                        const orderedLocations = partitionExistingThenNew(group.locations);
                        const locValues = orderedLocations.map(l => [
                            orgName,
                            l.name || l.Name || "",
                            l.id || l.Id || "",
                            l.active !== undefined ? (l.active ? "Active" : "Inactive") : "Active"
                        ]);
                        const locHighlightFlags = orderedLocations.map(isChangedRecord);
                        await writeRowsInBatches(context, sheet, "S", "V", currentRow, locValues, locHighlightFlags);
                    }

                    // Section 5 (X:AB) - Entities (Customers + Vendors), the
                    // section most likely to be large, batched the same way.
                    if (entityCount > 0) {
                        const orderedEntities = partitionExistingThenNew(group.entities);
                        const entityValues = orderedEntities.map(e => [
                            orgName,
                            e.name,
                            e.type,
                            e.id,
                            e.status
                        ]);
                        const entityHighlightFlags = orderedEntities.map(isChangedRecord);
                        await writeRowsInBatches(context, sheet, "X", "AB", currentRow, entityValues, entityHighlightFlags);
                    }

                    // Advance currentRow by maxRows + 2 (providing 2 empty row spaces between organizations!)
                    currentRow += maxRows + 2;
                }

                // Apply consistent font size and wrap text to prevent overlapping/truncation
                const dataRange = sheet.getRange("A2:AB10000");
                dataRange.format.font.size = 11;
                dataRange.format.wrapText = true;

                sheet.getRange("A:AB").format.columnWidth = 115;

                await context.sync();
            });
        },

        async clearMasterData() {
            await Excel.run(async (context) => {
                const masterSheet = context.workbook.worksheets.getItemOrNullObject("1.Master_Data");
                const inputSheet  = context.workbook.worksheets.getItemOrNullObject("2.Input");
                const sheets = context.workbook.worksheets;
                sheets.load("items/name");
                await context.sync();

                let otherSheetExists = false;
                for (let i = 0; i < sheets.items.length; i++) {
                    const name = sheets.items[i].name;
                    if (name !== "1.Master_Data" && name !== "2.Input") {
                        otherSheetExists = true;
                        sheets.items[i].activate();
                        break;
                    }
                }
                if (!otherSheetExists) {
                    context.workbook.worksheets.add("Sheet1").activate();
                }
                if (!masterSheet.isNullObject) masterSheet.delete();
                if (!inputSheet.isNullObject)  inputSheet.delete();
                await context.sync();
            });
        },

        async stampLastRefreshed(timestamp) {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItem("1.Master_Data");
                const cell = sheet.getRange("V1");
                cell.values = [[`Last Refreshed: ${timestamp}`]];
                cell.format.font.bold = true;
                cell.format.font.color = "white";
                await context.sync();
            });
        },

        /**
         * Scaffolds required sheets (1.Master_Data and 2.Input) and styles the header rows across 5 colored sections.
         * @param {string} provider - Active ERP provider
         */
        async setupWorkbookSheets(provider) {
            await Excel.run(async (context) => {
                let masterSheet = context.workbook.worksheets.getItemOrNullObject("1.Master_Data");
                let inputSheet  = context.workbook.worksheets.getItemOrNullObject("2.Input");
                await context.sync();

                if (masterSheet.isNullObject) masterSheet = context.workbook.worksheets.add("1.Master_Data");
                if (inputSheet.isNullObject)  inputSheet  = context.workbook.worksheets.add("2.Input");

                masterSheet.activate();
                await context.sync();

                const idLabel = provider === "quickbooks" ? "QBO" : "Xero";

                // Generate spreadsheet header columns from A1 to AB1
                const headerRange = masterSheet.getRange("A1:AB1");
                const headers = [
                    [
                        "Client ID", "Client Name", "", 
                        "Client ID", "Account Code", "Account Name", "Account Type", "Account Sub-Type", "Classification", "Fully Qualified Name", "Status", `${idLabel} Account Id`, "", 
                        "Client ID", "Class Name", `${idLabel} Class Id`, "Status", "", 
                        "Client ID", "Location Name", `${idLabel} Location Id`, "Status", "", 
                        "Client ID", "Entity Name", "Entity Type", `${idLabel} Entity Id`, "Status"
                    ]
                ];

                headerRange.clear();
                headerRange.values = headers;
                
                // Format Navy Blue Section 1 Header Range (A1:B1)
                const purpleRange1 = masterSheet.getRange("A1:B1");
                purpleRange1.format.fill.color = "#1B224C";
                purpleRange1.format.font.color = "white";
                purpleRange1.format.font.bold = true;
                purpleRange1.format.horizontalAlignment = "Center";
                purpleRange1.format.verticalAlignment = "Center";
                purpleRange1.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.Continuous;

                // Format Navy Blue Section 2 Header Range (D1:L1)
                const blueRange2 = masterSheet.getRange("D1:L1");
                blueRange2.format.fill.color = "#1F4E79";
                blueRange2.format.font.color = "white";
                blueRange2.format.font.bold = true;
                blueRange2.format.horizontalAlignment = "Center";
                blueRange2.format.verticalAlignment = "Center";
                blueRange2.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.Continuous;

                // Format Green Section 3 & 4 Header Ranges (N1:Q1, S1:V1)
                const greenRanges = ["N1:Q1", "S1:V1"];
                for (const range of greenRanges) {
                    const r = masterSheet.getRange(range);
                    r.format.fill.color = "#0F7546";
                    r.format.font.color = "white";
                    r.format.font.bold = true;
                    r.format.horizontalAlignment = "Center";
                    r.format.verticalAlignment = "Center";
                    r.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.Continuous;
                }

                // Format Dark Blue Section 5 Header Range (X1:AB1)
                const purpleRange5 = masterSheet.getRange("X1:AB1");
                purpleRange5.format.fill.color = "#1B224C";
                purpleRange5.format.font.color = "white";
                purpleRange5.format.font.bold = true;
                purpleRange5.format.horizontalAlignment = "Center";
                purpleRange5.format.verticalAlignment = "Center";
                purpleRange5.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.Continuous;

                // Configure spacer columns - clear formats
                const spacers = ["C:C", "M:M", "R:R", "W:W"];
                for (const spacer of spacers) {
                    const col = masterSheet.getRange(spacer);
                    col.clear("Formats");
                }

                headerRange.format.rowHeight = 28;
                headerRange.format.font.size = 11;
                headerRange.format.wrapText = true;
                
                // Set all columns in the range to column width 115
                masterSheet.getRange("A:AB").format.columnWidth = 115;
                
                masterSheet.freezePanes.unfreeze();
                masterSheet.getRange("A2").select();
                
                await context.sync();
            });
        }
    };

    // ============================================================
    // 5. AUTH SERVICE
    // ============================================================
    const AuthService = {
        /**
         * Called when a new-user completes payment inside the popup.
         * Receives full profile + subscription info from google_authed postMessage.
         * @param {string} email
         * @param {string} name
         * @param {string} provider
         * @param {string} subscriptionId
         * @param {string} plan
         */
        handleNewUserAuthed(email, name, provider, subscriptionId, plan, token) {
            AppState.userEmail        = email;
            AppState.userName         = name;
            AppState.userProvider     = provider;
            AppState.hasSubscription  = true;
            AppState.subscriptionId   = subscriptionId;
            AppState.subscriptionPlan = plan;

            // Persist the JWT token so all subsequent API calls are authenticated
            if (token) {
                AppState.jwtToken = token;
                AppState.sessionExpired = false; // fresh token — lift the request block
                localStorage.setItem("fa_jwt_token", token);
            }

            localStorage.setItem("fa_user_email",      email);
            localStorage.setItem("fa_user_name",       name);
            localStorage.setItem("fa_user_provider",   provider);
            this._persistSubscription();

            DashboardService.render();
            ViewRouter.show("Dashboard");
            DashboardService.showStatus("Login successful.", "success");
        },

        /**
         * Called when returning user signs in (popup closes immediately with google_profile)
         * and backend confirms their subscription.
         * @param {string} email
         * @param {string} name
         * @param {string} provider
         */
        async handleReturningUser(email, name, provider, token) {
            AppState.userEmail    = email;
            AppState.userName     = name;
            AppState.userProvider = provider;
            localStorage.setItem("fa_user_email",    email);
            localStorage.setItem("fa_user_name",     name);
            localStorage.setItem("fa_user_provider", provider);

            // Persist token from popup if provided (avoids a round-trip)
            if (token) {
                AppState.jwtToken = token;
                AppState.sessionExpired = false; // fresh token — lift the request block
                localStorage.setItem("fa_jwt_token", token);
            }

            ViewRouter.show("Loading");
            try {
                const result = await ApiService.checkSubscription(email);
                const userPlan = result.plan || (result.user && result.user.plan);
                if (result.hasSubscription || userPlan) {
                    AppState.hasSubscription  = true;
                    AppState.subscriptionId   = result.subscriptionId || AppState.subscriptionId;
                    AppState.subscriptionPlan = userPlan              || AppState.subscriptionPlan;
                    this._persistSubscription();
                    DashboardService.render();
                    ViewRouter.show("Dashboard");
                    DashboardService.showStatus("Login successful.", "success");
                } else {
                    // Plan is null or missing — show subscription page
                    ViewRouter.show("Plans");
                }
            } catch {
                ViewRouter.show("Plans");
                DashboardService.showStatus("Login failed. Please try again.", "error");
            }
        },

        /**
         * Opens a Google OAuth popup.
         * The popup now hosts the entire Plans → Payment → Success flow.
         * - For new users:      popup sends  { type: 'google_authed', email, name, subscriptionId, plan }
         * - For returning users (if future backend check skips popup): sends { type: 'google_profile', ... }
         * - On logout click:    popup sends  { type: 'google_cancelled' }
         */
        openGooglePopup() {
            const googleAuthUrl = `${ApiService.BASE}/api/auth/google/connect`;

            const msgHandler = (event) => {
                if (!event.data) return;
                let data = event.data;
                if (typeof data === "string") {
                    try { data = JSON.parse(data); } catch (_) {}
                }

                if (!data || !data.type) return;

                if (data.type === "google_authed") {
                    // New user completed payment inside popup
                    window.removeEventListener("message", msgHandler);
                    AuthService.handleNewUserAuthed(
                        data.email || "",
                        data.name  || data.email || "",
                        "google",
                        data.subscriptionId || "",
                        data.plan           || "Starter",
                        data.token          || ""
                    );
                } else if (data.type === "google_profile") {
                    // Returning user — popup closed immediately, check backend
                    window.removeEventListener("message", msgHandler);
                    AuthService.handleReturningUser(
                        data.email || "",
                        data.name  || data.email || "",
                        "google",
                        data.token || ""
                    );
                } else if (data.type === "google_cancelled") {
                    // User clicked logout in the popup
                    window.removeEventListener("message", msgHandler);
                    ViewRouter.show("Welcome");
                }
            };
            window.addEventListener("message", msgHandler);

            const popup = window.open(
                googleAuthUrl, "fa_google_auth",
                "width=640,height=840,top=40,left=80,toolbar=no,menubar=no"
            );

            if (!popup || popup.closed) {
                window.removeEventListener("message", msgHandler);
                DashboardService.showError("Popup was blocked. Please allow popups and try again.");
            } else {
                // Bring the popup to the front — without this it can open
                // behind the taskpane/main window in some browsers.
                popup.focus();
            }
        },

        /**
         * Opens a mock Microsoft OAuth flow.
         * In production, replace with real Microsoft MSAL / OAuth URL.
         */
        openMicrosoftPopup() {
            const mockMSUrl = `${ApiService.BASE}/api/microsoft/connect`;

            const msgHandler = (event) => {
                if (!event.data) return;
                let data = event.data;
                if (typeof data === "string") {
                    try { data = JSON.parse(data); } catch (_) {}
                }
                if (!data || !data.type) return;

                if (data.type === "microsoft_authed" || data.type === "ms_authed") {
                    window.removeEventListener("message", msgHandler);
                    AuthService.handleNewUserAuthed(
                        data.email || "",
                        data.name  || data.email || "",
                        "microsoft",
                        data.subscriptionId || "",
                        data.plan           || "Starter",
                        data.token          || ""
                    );
                } else if (data.type === "ms_profile" || data.type === "microsoft_profile") {
                    window.removeEventListener("message", msgHandler);
                    AuthService.handleReturningUser(
                        data.email || "",
                        data.name  || data.email || "",
                        "microsoft",
                        data.token || ""
                    );
                } else if (data.type === "ms_cancelled" || data.type === "google_cancelled") {
                    window.removeEventListener("message", msgHandler);
                    ViewRouter.show("Welcome");
                }
            };
            window.addEventListener("message", msgHandler);

            const popup = window.open(
                mockMSUrl, "fa_ms_auth",
                "width=640,height=800,top=60,left=80,toolbar=no,menubar=no"
            );

            if (!popup || popup.closed) {
                window.removeEventListener("message", msgHandler);
                DashboardService.showError("Popup was blocked. Please allow popups and try again.");
            } else {
                // Bring the popup to the front — without this it can open
                // behind the taskpane/main window in some browsers.
                popup.focus();
            }
        },

        _persistSubscription() {
            localStorage.setItem("fa_has_subscription",  String(AppState.hasSubscription));
            localStorage.setItem("fa_subscription_id",   AppState.subscriptionId   || "");
            const _planToSave = (AppState.subscriptionPlan && AppState.subscriptionPlan !== 'null' && AppState.subscriptionPlan !== 'undefined') ? AppState.subscriptionPlan : "";
            localStorage.setItem("fa_subscription_plan", _planToSave);
        },

        /**
         * Clears all auth + subscription state and returns to welcome screen.
         */
        logout() {
            ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data on logout: ", err));
            AppState.userEmail        = null;
            AppState.userName         = null;
            AppState.userProvider     = null;
            AppState.hasSubscription  = false;
            AppState.subscriptionId   = null;
            AppState.subscriptionPlan = null;
            AppState.erpConnected     = false;
            AppState.erpType          = null;
            AppState.erpOrgName       = null;
            AppState.erpConnectedDate = null;
            AppState.jwtToken         = null;  // Clear the JWT token on logout
            // Note: AppState.sessionExpired is deliberately NOT reset here —
            // it's only cleared once a fresh token is obtained via a
            // successful login (see checkSubscription / handleGoogleAuth /
            // handleReturningUser), so a burst of already-in-flight requests
            // failing right after logout can't re-trigger the redirect loop.

            hideBanner();

            [
                "fa_user_email", "fa_user_name", "fa_user_provider",
                "fa_has_subscription", "fa_subscription_id", "fa_subscription_plan",
                "fa_erp_connected", "fa_erp_type", "fa_erp_org", "fa_erp_date",
                "fa_current_company_id",
                "fa_last_view", "fa_jwt_token"
            ].forEach(k => localStorage.removeItem(k));

            try {
                ApiService.apiFetch("/api/auth/logout", { method: "POST" }).catch(() => {});
            } catch (_) {}

            ViewRouter.show("Welcome");
        }
    };

    // ============================================================
    // 6. CHECKOUT / PAYMENT SERVICE
    // ============================================================
    const CheckoutService = {
        /**
         * Initiates the mock hosted checkout flow.
         * In production: replace mock URL with Stripe/Razorpay checkout session URL.
         */
        openCheckout(plan, price, cycle) {
            AppState.pendingPlan  = plan;
            AppState.pendingPrice = price;
            AppState.pendingCycle = cycle;

            const tokenParam  = AppState.jwtToken ? `&token=${encodeURIComponent(AppState.jwtToken)}` : "";
            const checkoutUrl = `${ApiService.BASE}/api/payments/checkout?plan=${encodeURIComponent(plan)}&price=${price}&cycle=${encodeURIComponent(cycle)}&email=${encodeURIComponent(AppState.userEmail || "")}${tokenParam}`;

            const btnText    = document.getElementById("checkoutBtnText");
            const btnSpinner = document.getElementById("checkoutSpinner");
            if (btnText)    btnText.textContent = "Opening Secure Checkout...";
            if (btnSpinner) btnSpinner.classList.remove("hidden");

            const msgHandler = (event) => {
                if (!event.data) return;
                let data = event.data;
                if (typeof data === "string") {
                    try { data = JSON.parse(data); } catch (_) {}
                }
                if (data && (data.type === "payment_success" || data.type === "checkout_complete")) {
                    window.removeEventListener("message", msgHandler);
                    CheckoutService.handlePaymentSuccess(data);
                }
            };
            window.addEventListener("message", msgHandler);

            const popup = window.open(
                checkoutUrl, "fa_checkout",
                "width=540,height=700,top=60,left=80,toolbar=no,menubar=no"
            );

            if (btnText)    btnText.textContent = "Open Secure Checkout";
            if (btnSpinner) btnSpinner.classList.add("hidden");

            if (!popup || popup.closed) {
                window.removeEventListener("message", msgHandler);
                DashboardService.showError("Checkout popup was blocked. Please allow popups.");
            } else {
                popup.focus();
            }
        },

        /**
         * Handles successful payment message from checkout popup.
         * In production, backend verifies and returns subscription details.
         */
        handlePaymentSuccess(data) {
            const subId = data.subscriptionId || ("FA-SUB-" + Math.floor(100000 + Math.random() * 900000));
            const plan  = data.plan           || AppState.pendingPlan  || "Professional";

            AppState.hasSubscription  = true;
            AppState.subscriptionId   = subId;
            AppState.subscriptionPlan = plan;
            AuthService._persistSubscription();

            // Show success screen
            const idEl   = document.getElementById("successSubId");
            const planEl = document.getElementById("successPlanName");
            if (idEl)   idEl.textContent   = subId;
            if (planEl) planEl.textContent  = plan;

            ViewRouter.show("Success");
            DashboardService.showStatus("Payment successful.", "success", `Subscribed to the ${plan} plan.`);
        },

        /**
         * Manually verifies a payment when user clicks "Verify my payment".
         * In production: calls GET /api/payments/verify?email=...
         */
        async verifyPayment() {
            ViewRouter.show("Loading");
            try {
                const res    = await ApiService.checkSubscription(AppState.userEmail);
                if (res.hasSubscription) {
                    AppState.hasSubscription  = true;
                    AppState.subscriptionId   = res.subscriptionId || AppState.subscriptionId;
                    AppState.subscriptionPlan = res.plan           || AppState.pendingPlan;
                    AuthService._persistSubscription();

                    const idEl   = document.getElementById("successSubId");
                    const planEl = document.getElementById("successPlanName");
                    if (idEl)   idEl.textContent   = AppState.subscriptionId;
                    if (planEl) planEl.textContent  = AppState.subscriptionPlan;
                    ViewRouter.show("Success");
                    DashboardService.showStatus("Payment successful.", "success", `Subscribed to the ${AppState.subscriptionPlan} plan.`);
                } else {
                    ViewRouter.show("Payment");
                    DashboardService.showStatus("Payment verification failed.", "error", "We couldn't find an active subscription yet. Please try again.");
                }
            } catch {
                ViewRouter.show("Payment");
                DashboardService.showStatus("Payment verification failed.", "error", "Please try again.");
            }
        }
    };

    // ============================================================
    // 7. DASHBOARD SERVICE
    // ============================================================
    // ============================================================
    // 7. DASHBOARD SERVICE
    // ============================================================
    const DashboardService = {
        /**
         * Renders and populates all dashboard UI elements based on AppState.
         */
        render() {
            const name  = AppState.userName  || AppState.userEmail || "User";
            const first = name.split(" ")[0];
            const initial = name.charAt(0).toUpperCase();

            // Set avatar initials
            const avatars = ["dashHeaderAvatarBtn1", "dashHeaderAvatarBtn2", "dashHeaderAvatarBtn3", "dropdownAvatar", "blockAvatar"];
            avatars.forEach(id => {
                const el = document.getElementById(id);
                if (el) el.textContent = initial;
            });

            // Render details in disconnected state header
            const welcomeEl = document.getElementById("dashWelcome");
            const badgeEl   = document.getElementById("dashPlanBadge");
            const subIdEl   = document.getElementById("dashSubId");

            if (welcomeEl) welcomeEl.textContent = `Welcome, ${first}!`;
            if (badgeEl)   badgeEl.textContent = ((AppState.subscriptionPlan && AppState.subscriptionPlan !== 'null' ? AppState.subscriptionPlan : "Basic") + " Plan");
            if (subIdEl)   subIdEl.textContent = AppState.subscriptionId || "—";
            
            // Render details in dropdown and blocks
            if (document.getElementById("dropdownUserName")) document.getElementById("dropdownUserName").textContent = name;
            if (document.getElementById("dropdownUserEmail")) document.getElementById("dropdownUserEmail").textContent = AppState.userEmail;
            if (document.getElementById("blockUserName")) document.getElementById("blockUserName").textContent = name;
            if (document.getElementById("blockUserEmail")) document.getElementById("blockUserEmail").textContent = AppState.userEmail;
            
            if (document.getElementById("blockSubId")) document.getElementById("blockSubId").textContent = AppState.subscriptionId || "—";
            const _safePlan = (AppState.subscriptionPlan && AppState.subscriptionPlan !== 'null') ? AppState.subscriptionPlan : "Basic";
            if (document.getElementById("blockPlanName")) document.getElementById("blockPlanName").textContent = _safePlan + " Plan";
            if (document.getElementById("blockPlanTitle")) document.getElementById("blockPlanTitle").textContent = _safePlan + " Plan";
            if (document.getElementById("blockPlanType")) document.getElementById("blockPlanType").textContent = _safePlan + " Plan";

            // Render correct state sections
            this.renderERPSection();

            // Set connected values if connected
            if (AppState.erpConnected && AppState.erpType) {
                AppState.currentProvider = AppState.erpType;
                this.renderERPConsole();
            }
        },

        /**
         * Shows the provider-selected "not connected" state (Image 4).
         * Called when user clicks QuickBooks or Xero card from the connect screen.
         * @param {"quickbooks"|"xero"} provider
         */
        showProviderSelected(provider) {
            AppState.currentProvider = provider;
            const isQB  = provider === "quickbooks";
            const pName = isQB ? "QuickBooks" : "Xero";

            // Hide disconnected + connected, show provider-selected
            const discSection = document.getElementById("dashDisconnected");
            const provSection = document.getElementById("dashProviderSelected");
            const connSection = document.getElementById("dashConnected");
            const connectingSection = document.getElementById("dashConnecting");
            if (discSection) discSection.style.display = "none";
            if (provSection) provSection.style.display = "flex";
            if (connSection) connSection.style.display = "none";
            if (connectingSection) connectingSection.style.display = "none";

            // Update the plan badge
            const planBadge = document.getElementById("providerPlanBadge");
            if (planBadge) {
                planBadge.textContent = isQB ? "QBO PRO" : "XERO PRO";
            }

            // Update the connect button
            const connectBtn = document.getElementById("btnConnectProvider");
            if (connectBtn) connectBtn.textContent = `Connect ${pName}`;

            // Update pull button label
            const pullLabel = document.getElementById("pullBtnProvLabel");
            if (pullLabel) pullLabel.textContent = isQB ? "QBO" : "Xero";

            // Update progress step labels
            const step1Label = document.getElementById("provStep1Label");
            const step3Label = document.getElementById("provStep3Label");
            if (step1Label) step1Label.textContent = `Connect to ${pName}`;
            if (step3Label) step3Label.textContent = `Pull Master Data`;

            // Active provider just changed — keep the badge/drawer, log
            // console, and step indicators in sync (all re-scoped to this
            // provider + whatever company is active for it).
            NotificationService.refreshForContext();
            this.renderActiveLogConsole();
            this.applyStepState();
        },

        showConnecting(provider) {
            AppState.currentProvider = provider;
            const pName = provider === "quickbooks" ? "QuickBooks" : "Xero";

            const discSection = document.getElementById("dashDisconnected");
            const provSection = document.getElementById("dashProviderSelected");
            const connSection = document.getElementById("dashConnected");
            const connectingSection = document.getElementById("dashConnecting");

            if (discSection) discSection.style.display = "none";
            if (provSection) provSection.style.display = "none";
            if (connSection) connSection.style.display = "none";
            if (connectingSection) {
                connectingSection.style.display = "flex";
                const textEl = document.getElementById("connectingText");
                if (textEl) textEl.textContent = `Connecting with ${pName}...`;
            }

            // Active provider just changed — keep the badge/drawer, log
            // console, and step indicators in sync.
            NotificationService.refreshForContext();
            this.renderActiveLogConsole();
            this.applyStepState();
        },

        /**
         * Shows the correct ERP section based on state:
         * - disconnected: connect cards (Image 2)
         * - provider-selected: not connected (Image 4)
         * - connected: fully connected (Image 3)
         */
        renderERPSection() {
            const discSection = document.getElementById("dashDisconnected");
            const provSection = document.getElementById("dashProviderSelected");
            const connSection = document.getElementById("dashConnected");
            const connectingSection = document.getElementById("dashConnecting");

            ApiService.apiFetch("/api/connections?mail=" + encodeURIComponent(AppState.userEmail || ""))
                .then(r => r.json())
                .then(conns => {
                    const dropdown = document.getElementById("companySelectDropdown");
                    const companyListEl = document.getElementById("companyList");
                    const footerEl = document.getElementById("companyListFooter");
                    const modalListEl = document.getElementById("modalCompanyList");

                    const activeConns = conns.filter(c => c.status !== 'Disconnected');
                    if (activeConns.length > 0) {
                        AppState.forceWelcome = false;
                    }
                    if (!conns || conns.length === 0 || AppState.forceWelcome) {
                        AppState.erpConnected = false;
                        if (discSection) {
                            discSection.style.display = "flex";
                            discSection.style.flexDirection = "column";
                            discSection.style.height = "100%";
                        }
                        if (provSection) provSection.style.display = "none";
                        if (connSection) connSection.style.display = "none";
                        if (connectingSection) connectingSection.style.display = "none";

                        // Update QB card button label
                        const hasQB   = (conns || []).some(c => (c.platform || "").toLowerCase() === "quickbooks");
                        const hasXero = (conns || []).some(c => (c.platform || "").toLowerCase() === "xero");
                        const qbBtn   = document.querySelector("#btnConnectQB .btn-connect-full");
                        const xeroBtn = document.querySelector("#btnConnectXero .btn-connect-full");
                        if (qbBtn)   qbBtn.innerHTML   = hasQB   ? "▶ Open QuickBooks Dashboard →" : `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path></svg> Connect QuickBooks →`;
                        if (xeroBtn) xeroBtn.innerHTML = hasXero ? "▶ Open Xero Dashboard →"       : `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path></svg> Connect Xero →`;

                        // No active provider anymore — recompute the badge/drawer,
                        // log console, and step indicators so nothing left over
                        // from a previous provider/company lingers.
                        NotificationService.refreshForContext();
                        this.renderActiveLogConsole();
                        this.applyStepState();
                        return;
                    }

                    // We have connections!
                    AppState.erpConnected = true;
                    if (discSection) discSection.style.display = "none";
                    if (provSection) provSection.style.display = "none";
                    if (connSection) connSection.style.display = "flex";
                    if (connectingSection) connectingSection.style.display = "none";

                    // Determine the current platform first (respect AppState.currentProvider if already set)
                    let resolvedProvider = AppState.currentProvider || null;
                    if (!resolvedProvider) {
                        // Derive from the currently tracked companyId if possible
                        const existingConn = activeConns.find(c => c.companyId === AppState.currentCompanyId) || conns.find(c => c.companyId === AppState.currentCompanyId) || conns[0];
                        resolvedProvider = existingConn ? (existingConn.platform || "quickbooks").toLowerCase() : "quickbooks";
                    }
                    AppState.currentProvider = resolvedProvider;
                    AppState.erpType = resolvedProvider;

                    // Get active connections scoped to the current platform
                    const platformActiveConns = activeConns.filter(c => (c.platform || "quickbooks").toLowerCase() === resolvedProvider);
                    const fallbackConns = platformActiveConns.length > 0 ? platformActiveConns : conns.filter(c => (c.platform || "quickbooks").toLowerCase() === resolvedProvider);

                    // Ensure current active company ID is valid for this platform
                    const isCurrentValid = AppState.currentCompanyId && platformActiveConns.some(ac => ac.companyId === AppState.currentCompanyId);
                    if (!isCurrentValid && fallbackConns.length > 0) {
                        AppState.currentCompanyId = fallbackConns[0].companyId;
                        // Update provider if we fell back to a different platform
                        const fb = fallbackConns[0];
                        AppState.currentProvider = (fb.platform || "quickbooks").toLowerCase();
                        AppState.erpType = AppState.currentProvider;
                    }

                    // Determine provider
                    const isXero = (AppState.currentProvider || "quickbooks").toLowerCase() === "xero";

                    // Handle Section Visibility (Xero Multi-select vs QB Single Active Company)
                    const selectXeroCard = document.getElementById("selectXeroCompaniesCard");
                    const activeCompanyCard = document.getElementById("activeCompanyCard");

                    if (isXero) {
                        if (selectXeroCard) selectXeroCard.style.display = "block";
                        if (activeCompanyCard) activeCompanyCard.style.display = "none";

                        // Populate Xero Multi-Select Checklist Card
                        const currentPlan = (AppState.subscriptionPlan && AppState.subscriptionPlan !== 'null') ? AppState.subscriptionPlan : "Basic";
                        const maxAllowed = currentPlan === "Basic" ? 1 : (currentPlan === "Standard" ? 3 : 10);
                        const xeroConns = conns.filter(c => (c.platform || "").toLowerCase() === "xero");
                        
                        let xeroSelected = new Set(xeroConns.filter(c => c.status !== 'Disconnected').map(c => c.companyId));

                        const listEl = document.getElementById("xeroCompanyList");
                        const selCountEl = document.getElementById("xeroSelCount");
                        const maxCountEl = document.getElementById("xeroMaxCount");
                        const maxAllowedEl = document.getElementById("xeroMaxAllowed");
                        const planBadgeEl = document.getElementById("xeroPlanBadge");
                        const warningEl = document.getElementById("xeroLimitWarning");
                        const confirmBtn = document.getElementById("btnConfirmXeroCompanies");

                        if (maxCountEl) maxCountEl.textContent = maxAllowed;
                        if (maxAllowedEl) maxAllowedEl.textContent = maxAllowed;
                        if (planBadgeEl) planBadgeEl.textContent = `${currentPlan.toUpperCase()} (${maxAllowed}) PLAN`;

                        const updateXeroUI = () => {
                            if (selCountEl) selCountEl.textContent = xeroSelected.size;
                            if (confirmBtn) confirmBtn.disabled = xeroSelected.size === 0;
                            if (warningEl) warningEl.style.display = xeroSelected.size >= maxAllowed ? "inline" : "none";
                        };

                        if (listEl) {
                            listEl.innerHTML = "";
                            xeroConns.forEach(c => {
                                const isChecked = xeroSelected.has(c.companyId);
                                const row = document.createElement("div");
                                row.className = `xero-company-row ${isChecked ? "selected" : ""}`;
                                row.id = "xero_row_" + c.companyId;
                                row.innerHTML = `
                                    <input type="checkbox" class="xero-company-cb" id="xero_cb_${c.companyId}" value="${c.companyId}" ${isChecked ? "checked" : ""} />
                                    <div class="xero-company-icon">xero</div>
                                    <div class="xero-company-info">
                                        <div class="xero-company-name">${c.companyName || "Xero Organisation"}</div>
                                        <div class="xero-company-id">Tenant: ${c.companyId || "—"}</div>
                                    </div>
                                `;

                                const cb = row.querySelector(".xero-company-cb");

                                row.addEventListener("click", (e) => {
                                    if (e.target === cb) return;
                                    if (cb.checked) {
                                        cb.checked = false;
                                    } else {
                                        if (xeroSelected.size >= maxAllowed && !xeroSelected.has(c.companyId)) return;
                                        cb.checked = true;
                                    }
                                    cb.dispatchEvent(new Event("change"));
                                });

                                cb.addEventListener("change", () => {
                                    if (cb.checked) {
                                        if (xeroSelected.size >= maxAllowed) {
                                            cb.checked = false;
                                            return;
                                        }
                                        xeroSelected.add(c.companyId);
                                        row.classList.add("selected");
                                    } else {
                                        xeroSelected.delete(c.companyId);
                                        row.classList.remove("selected");
                                    }
                                    updateXeroUI();
                                });

                                listEl.appendChild(row);
                            });
                        }
                        updateXeroUI();

                        if (confirmBtn) {
                            const newConfirmBtn = confirmBtn.cloneNode(true);
                            confirmBtn.parentNode.replaceChild(newConfirmBtn, confirmBtn);

                            newConfirmBtn.addEventListener("click", async () => {
                                if (xeroSelected.size === 0) return;
                                newConfirmBtn.disabled = true;
                                newConfirmBtn.innerHTML = '<span class="spinner" style="display:inline-block;width:12px;height:12px;border:2px solid rgba(255,255,255,0.4);border-top-color:#fff;border-radius:50%;animation:spin 0.7s linear infinite;margin-right:6px;vertical-align:middle;"></span> Saving...';

                                try {
                                    for (const c of xeroConns) {
                                        if (xeroSelected.has(c.companyId)) {
                                            await ApiService.apiFetch(`/api/connections/${c.companyId}/activate`, { method: "POST" });
                                        } else {
                                            await ApiService.apiFetch(`/api/connections/${c.companyId}`, { method: "DELETE" });
                                        }
                                    }

                                    const activeId = Array.from(xeroSelected)[0];
                                    if (activeId) AppState.currentCompanyId = activeId;

                                    DashboardService.showStatus("Xero companies saved successfully.", "success", null, "xero");
                                    DashboardService.renderERPSection();
                                } catch (err) {
                                    console.error("Error saving Xero companies:", err);
                                    DashboardService.showStatus("Failed to save Xero companies.", "error", null, "xero");
                                    newConfirmBtn.disabled = false;
                                    newConfirmBtn.textContent = "Connect Selected Companies";
                                }
                            });
                        }
                    } else {
                        if (selectXeroCard) selectXeroCard.style.display = "none";
                        if (activeCompanyCard) activeCompanyCard.style.display = "block";
                    }

                    // Populate Active Company Dropdown (platform-filtered)
                    if (dropdown) {
                        dropdown.innerHTML = "";
                        const platformDropdownConns = activeConns.filter(c => (c.platform || "quickbooks").toLowerCase() === (AppState.currentProvider || "quickbooks"));
                        platformDropdownConns.forEach(c => {
                            const opt = document.createElement("option");
                            opt.value = c.companyId;
                            opt.dataset.platform = (c.platform || "QuickBooks").toLowerCase();
                            opt.textContent = `${c.companyName || "Company"}`;
                            if (c.companyId === AppState.currentCompanyId) opt.selected = true;
                            dropdown.appendChild(opt);
                        });
                    }

                    // Set active company details
                    const activeConn = activeConns.find(c => c.companyId === AppState.currentCompanyId) || activeConns[0];
                    if (activeConn) {
                        AppState.currentCompanyId = activeConn.companyId;
                        AppState.currentProvider = (activeConn.platform || "quickbooks").toLowerCase();
                        AppState.erpType = AppState.currentProvider;

                        // Persist the resolved active company so a refresh
                        // (which re-creates AppState from scratch) restores
                        // the same company instead of always falling back
                        // to the first one in the list. This is the single
                        // point every switch path (radio button, dropdown)
                        // routes back through via renderERPSection(), so it
                        // covers both QuickBooks and Xero uniformly.
                        localStorage.setItem("fa_current_company_id", activeConn.companyId || "");

                        const activeRealmEl = document.getElementById("activeRealmId");
                        if (activeRealmEl) activeRealmEl.textContent = activeConn.companyId || "—";
                        const realmEl = document.getElementById("connRealmId");
                        if (realmEl) realmEl.textContent = activeConn.companyId || "—";
                    }

                    // Platform-scoped connections (only show same platform as current active company)
                    const currentPlatform = AppState.currentProvider || "quickbooks";
                    const platformConns = conns.filter(c => (c.platform || "quickbooks").toLowerCase() === currentPlatform);

                    // Populate Company Management List (platform-filtered)
                    if (companyListEl) {
                        companyListEl.innerHTML = "";
                        platformConns.forEach(c => {
                            const isActive = c.companyId === AppState.currentCompanyId;
                            const isXero = (c.platform || "").toLowerCase() === "xero";
                            const isDisconnected = c.status === 'Disconnected';
                            const item = document.createElement("div");
                            item.className = `fa-company-item ${isActive ? "active-company" : ""} ${isDisconnected ? "disconnected-company" : ""}`;
                            item.dataset.companyId = c.companyId;
                            
                            let badgeOrBtn = "";
                            if (isActive && !isDisconnected) {
                                badgeOrBtn = '<span class="fa-badge-active">ACTIVE</span>';
                            } else if (isDisconnected) {
                                badgeOrBtn = '<button class="fa-btn-reconnect">Reconnect</button>';
                            } else {
                                badgeOrBtn = '<button class="fa-btn-switch">Switch</button>';
                            }

                            item.innerHTML = `
                                <input type="radio" name="companyRadio" class="fa-company-radio" ${isActive && !isDisconnected ? "checked" : ""} />
                                <div class="fa-company-icon ${isXero ? 'xero-company-icon' : ''}">${isXero ? 'xero' : 'qb'}</div>
                                <div class="fa-company-info">
                                    <div class="fa-company-name">${c.companyName || (isXero ? "Xero Organisation" : "QuickBooks Company")}</div>
                                    <div class="fa-company-tag">Last Sync: ${this.formatRelativeTime(c.lastSyncedAt, c.status)}</div>
                                </div>
                                <div class="fa-company-actions">
                                    ${badgeOrBtn}
                                    <button class="fa-btn-dots" title="More options">⋮</button>
                                </div>
                            `;

                            // Click radio or card row to make active (ignore for disconnected companies)
                            item.addEventListener("click", (e) => {
                                if (e.target.classList.contains("fa-btn-dots")) {
                                    e.stopPropagation();
                                    this.showContextMenu(e.target, c);
                                    return;
                                }
                                if (e.target.classList.contains("fa-btn-reconnect")) {
                                    e.stopPropagation();
                                    const reconnectPlatform = (c.platform || "quickbooks").toLowerCase();
                                    this.showStatus("Launching re-authorization...", "success", null, reconnectPlatform);
                                    this.launchERPOAuth(reconnectPlatform);
                                    return;
                                }
                                if (isDisconnected) {
                                    // Do not activate disconnected companies on box click
                                    return;
                                }
                                this.switchActiveCompany(c.companyId, platformConns);
                            });

                            companyListEl.appendChild(item);
                        });
                    }

                    if (footerEl) {
                        const platformLabel = currentPlatform === "xero" ? "Xero" : "QuickBooks";
                        footerEl.textContent = `Showing ${platformConns.length} ${platformLabel} ${platformConns.length === 1 ? "company" : "companies"}`;
                    }

                    // Populate Subscription Stats (filtered by current active platform: quickbooks or xero)
                    const currentPlan = (AppState.subscriptionPlan && AppState.subscriptionPlan !== 'null') ? AppState.subscriptionPlan : "Basic";
                    const maxAllowed = currentPlan === "Basic" ? 1 : (currentPlan === "Standard" ? 3 : 10);
                    
                    const connectedCount = platformConns.length;
                    const remaining = Math.max(0, maxAllowed - connectedCount);

                    if (document.getElementById("subInfoPlan")) document.getElementById("subInfoPlan").textContent = currentPlan;
                    if (document.getElementById("subInfoConnected")) document.getElementById("subInfoConnected").textContent = `${connectedCount} / ${maxAllowed}`;
                    if (document.getElementById("subInfoRemaining")) document.getElementById("subInfoRemaining").textContent = String(remaining);

                    // Update Tier Badge in Header
                    const tierBadge = document.getElementById("connTierBadge");
                    if (tierBadge) tierBadge.textContent = `${currentPlan.toUpperCase()} PLAN`;

                    // Update dynamic button label for Pull Master Data
                    const isQB = AppState.currentProvider === "quickbooks";
                    const platformDisplayName = isQB ? "QuickBooks" : "Xero";
                    const pullLabel = document.getElementById("pullBtnLabel");
                    if (pullLabel) pullLabel.textContent = isQB ? "QBO" : "Xero";

                    // Update "Company Management" section title to reflect platform
                    const sectionTitle = document.querySelector(".fa-section-title");
                    if (sectionTitle) sectionTitle.textContent = `${platformDisplayName} Companies`;

                    // Update header status realm label (kept beside the ID
                    // itself, not as a separate label elsewhere in the row)
                    const connStatus = document.querySelector(".fa-conn-status");
                    if (connStatus && activeConn) {
                        const idLabel = isQB ? "Realm ID" : "Tenant ID";
                        connStatus.innerHTML = `<span class="fa-realm-label">${idLabel}:</span> <span id="connRealmId">${activeConn.companyId || "—"}</span>`;
                    }

                    // Update Disconnect button label
                    const disconnectBtn = document.getElementById("btnDisconnectERP");
                    if (disconnectBtn) disconnectBtn.textContent = `Disconnect ${platformDisplayName}`;

                    // Show console for correct provider
                    const qbConsole   = document.getElementById("qbConsole");
                    const xeroConsole = document.getElementById("xeroConsole");
                    if (qbConsole)   qbConsole.style.display   = isQB ? "flex" : "none";
                    if (xeroConsole) xeroConsole.style.display  = isQB ? "none" : "flex";

                    // The active provider/company may have just changed
                    // (switch/resume/connect) — recompute the badge/drawer,
                    // log console, and step indicators so all three stay
                    // scoped to exactly what's active now, never a leftover
                    // from before.
                    NotificationService.refreshForContext();
                    this.renderActiveLogConsole();
                    this.applyStepState();
                })
                .catch(() => {
                    // Fallback to offline/disconnected view
                    if (discSection) {
                        discSection.style.display = "flex";
                        discSection.style.flexDirection = "column";
                        discSection.style.height = "100%";
                    }
                    if (provSection) provSection.style.display = "none";
                    if (connSection) connSection.style.display = "none";
                    if (connectingSection) connectingSection.style.display = "none";
                });
        },

        formatRelativeTime(dateInput, status) {
            if (status === 'Disconnected') return "Disconnected";
            // A freshly connected company starts with status 'Not Synced'
            // and no lastSyncedAt — show that literally instead of a vaguer
            // "not synced yet". Once the first Master Data Pull succeeds,
            // the backend flips status to 'Active' and stamps lastSyncedAt,
            // so this falls through to the relative-time formatting below
            // (which reports "Just now" immediately after that pull).
            if (status === 'Not Synced' || !dateInput) return "Not Synced";
            const date = new Date(dateInput);
            if (isNaN(date.getTime())) return "Not Synced";
            const now = new Date();
            const diffMs = now - date;
            if (diffMs < 0) return "Just now";
            const diffSec = Math.floor(diffMs / 1000);
            if (diffSec < 45) return "Just now";
            const diffMin = Math.floor(diffSec / 60);
            if (diffMin < 60) {
                return diffMin === 1 ? "1 minute ago" : `${diffMin} minutes ago`;
            }
            const diffHr = Math.floor(diffMin / 60);
            if (diffHr < 24) {
                return diffHr === 1 ? "1 hour ago" : `${diffHr} hours ago`;
            }
            const diffDays = Math.floor(diffHr / 24);
            if (diffDays < 30) {
                return diffDays === 1 ? "1 day ago" : `${diffDays} days ago`;
            }
            return date.toLocaleDateString();
        },

        switchActiveCompany(companyId, conns) {
            AppState.currentCompanyId = companyId;

            // Set the provider immediately from the target company so renderERPSection shows correct platform
            const targetConn = conns.find(c => c.companyId === companyId);
            const targetPlatform = targetConn ? (targetConn.platform || "quickbooks").toLowerCase() : null;
            if (targetConn) {
                AppState.currentProvider = targetPlatform;
                AppState.erpType = AppState.currentProvider;
            }

            ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data: ", err));

            // Toast only fires once the backend confirms the switch — not
            // when the click merely starts the request.
            ApiService.apiFetch(`/api/connections/${companyId}/activate`, { method: "POST" })
                .then(() => {
                    this.renderERPSection();
                    if (targetConn) {
                        this.addLog(`Switched active company to: ${targetConn.companyName}`);
                        this.showStatus(`Active company updated to ${targetConn.companyName}`, "success", null, targetPlatform);
                    }
                })
                .catch(err => {
                    console.error("Error activating company:", err);
                    this.renderERPSection();
                    this.showStatus("Failed to switch active company.", "error", null, targetPlatform);
                });
        },

        showContextMenu(targetBtn, company) {
            const menu = document.getElementById("companyContextMenu");
            if (!menu) return;
            const rect = targetBtn.getBoundingClientRect();
            menu.style.top = `${rect.bottom + 4}px`;
            menu.style.left = `${rect.left - 100}px`;
            menu.style.display = "block";

            const editBtn = document.getElementById("ctxEdit");
            const disconnectBtn = document.getElementById("ctxDisconnect");

            if (editBtn) {
                editBtn.onclick = () => {
                    menu.style.display = "none";
                    this.showRenameModal(company);
                };
            }
            if (disconnectBtn) {
                disconnectBtn.onclick = async () => {
                    menu.style.display = "none";
                    const companyPlatform = (company.platform || "quickbooks").toLowerCase();
                    this.showStatus(`Disconnecting ${company.companyName}...`, "success", null, companyPlatform);
                    try {
                        await ApiService.apiFetch(`/api/connections/${company.companyId}`, { method: "DELETE" });
                        // If we disconnected the currently active company, reset it so a new one is picked
                        if (AppState.currentCompanyId === company.companyId) {
                            AppState.currentCompanyId = null;
                        }
                        try { await ExcelService.clearMasterData(); } catch (_) {}
                        this.showStatus("Company disconnected.", "success", null, companyPlatform);
                        this.renderERPSection();
                    } catch (_) {
                        this.showStatus("Failed to disconnect company.", "error", null, companyPlatform);
                    }
                };
            }
        },

        showRenameModal(company) {
            const modal = document.getElementById("renameCompanyModal");
            const input = document.getElementById("renameCompanyInput");
            const closeBtn = document.getElementById("btnCloseRenameCompany");
            const cancelBtn = document.getElementById("btnCancelRenameCompany");
            const confirmBtn = document.getElementById("btnConfirmRenameCompany");

            if (!modal || !input) return;
            input.value = company.companyName || "";
            modal.style.display = "flex";
            input.focus();

            const closeModal = () => {
                modal.style.display = "none";
            };

            if (closeBtn) closeBtn.onclick = closeModal;
            if (cancelBtn) cancelBtn.onclick = closeModal;

            if (confirmBtn) {
                confirmBtn.onclick = async () => {
                    const newName = input.value.trim();
                    if (!newName) return;
                    const companyPlatform = (company.platform || "quickbooks").toLowerCase();
                    confirmBtn.disabled = true;
                    try {
                        const res = await ApiService.apiFetch(`/api/connections/${company.companyId}/rename`, {
                            method: "PATCH",
                            headers: { "Content-Type": "application/json" },
                            body: JSON.stringify({ companyName: newName })
                        });
                        if (res.ok) {
                            company.companyName = newName;
                            this.showStatus("Company renamed successfully.", "success", null, companyPlatform);
                            closeModal();
                            this.renderERPSection();
                        } else {
                            this.showStatus("Failed to rename company.", "error", null, companyPlatform);
                        }
                    } catch (err) {
                        console.error("Rename error:", err);
                        this.showStatus("Failed to rename company.", "error", null, companyPlatform);
                    } finally {
                        confirmBtn.disabled = false;
                    }
                };
            }
        },

        /**
         * Updates the progress step markers for the ERP console.
         */
        renderERPConsole() {
            const stepId = AppState.erpType === "quickbooks" ? "stepConnect" : "xeroStepConnect";
            const stepEl = document.getElementById(stepId);
            if (stepEl) stepEl.classList.add("complete");
        },

        // Console log history persists across taskpane refreshes (the
        // add-in's DOM/JS state is rebuilt from scratch on every reload, so
        // without this the QuickBooks/Xero log consoles would always come
        // back empty). Capped so localStorage can't grow unbounded.
        //
        // Every entry is tagged with both the provider AND the company that
        // was active when it was logged. The console only ever renders
        // entries matching the CURRENT provider + active company — this is
        // what stops a log line like "Pulling master data..." from another
        // company (or an earlier, unrelated session) from ever appearing
        // as if it just happened. Each entry only shows up once the real
        // action it describes has actually run for the company you're
        // currently looking at.
        LOG_STORAGE_KEY: "fa_console_logs",
        MAX_LOG_ENTRIES: 300,

        /** @returns {Array<{provider:('quickbooks'|'xero'), companyId:(string|null), message:string, timestamp:string}>} */
        _loadLogs() {
            try {
                const raw = localStorage.getItem(this.LOG_STORAGE_KEY);
                const parsed = raw ? JSON.parse(raw) : [];
                return Array.isArray(parsed) ? parsed : [];
            } catch (_) {
                return [];
            }
        },

        _saveLogs(list) {
            try {
                localStorage.setItem(this.LOG_STORAGE_KEY, JSON.stringify(list));
            } catch (_) {
                // Storage full/unavailable — logs just won't persist across
                // a refresh, but the app keeps working.
            }
        },

        /**
         * Renders a single stored/new entry as a log line and appends it to
         * the given console element, without re-stamping the time — the
         * original timestamp is preserved exactly as logged.
         * @param {HTMLElement} log
         * @param {string} message
         * @param {string} timestampIso
         */
        _appendLogLine(log, message, timestampIso) {
            const line = document.createElement("div");
            line.className = "log-line";
            if (message.toLowerCase().includes("error")) {
        line.style.color = "#ef4444"; // Red color for errors
    }
            const timeLabel = new Date(timestampIso).toLocaleTimeString();
            line.textContent = `[${timeLabel}] ${message}`;
            log.appendChild(line);
        },

        /**
         * Re-renders the visible QuickBooks/Xero console from persisted
         * history, filtered strictly to the current provider AND the
         * currently active company (AppState.currentCompanyId). Call this
         * any time the active provider or company changes (connect,
         * switch, resume, disconnect, dropdown/modal change) as well as on
         * app init — it's the single source of truth for what the console
         * shows, so a company that's never had Setup/Pull run always
         * starts with a clean console, never another company's history.
         */
        renderActiveLogConsole() {
            const provider = AppState.currentProvider === "quickbooks" ? "quickbooks" : "xero";
            const logId = provider === "quickbooks" ? "qbLog" : "xeroLog";
            const log = document.getElementById(logId);
            if (!log) return;

            const companyId = AppState.currentCompanyId || null;
            const entries = this._loadLogs().filter(entry =>
                entry.provider === provider && (entry.companyId || null) === companyId
            );

            log.innerHTML = "";
            entries.forEach(entry => this._appendLogLine(log, entry.message, entry.timestamp));
            log.scrollTop = log.scrollHeight;
        },

        /**
         * Adds a log line for an action that has actually just executed,
         * persists it tagged to the current provider + active company, and
         * re-renders the console from that persisted history. Because the
         * console always re-derives its content from storage (filtered to
         * the exact provider/company in view), a log entry can only ever
         * appear after the real action it describes has run — nothing is
         * fabricated or carried over from a different company.
         * @param {string} message
         */
        addLog(message) {
            const provider = AppState.currentProvider === "quickbooks" ? "quickbooks" : "xero";
            const companyId = AppState.currentCompanyId || null;
            const timestamp = new Date().toISOString();

            const list = this._loadLogs();
            list.push({ provider, companyId, message: String(message), timestamp });
            if (list.length > this.MAX_LOG_ENTRIES) {
                list.splice(0, list.length - this.MAX_LOG_ENTRIES);
            }
            this._saveLogs(list);

            this.renderActiveLogConsole();
        },

        /**
         * Pops an immediate top-right toast and records the same
         * notification in the Bell Notification Center. Transient
         * "in progress" messages (e.g. "Pulling data...") are not turned
         * into a toast/notification — only terminal outcomes are, so the
         * user isn't shown a toast for every intermediate step. No banner
         * is rendered inline; status updates live in the log console only.
         *
         * `provider` tags a QuickBooks/Xero-specific outcome so it's only
         * ever shown while the user is actively on that same ERP — pass
         * "quickbooks" or "xero" for provider-scoped actions (connect,
         * setup sheets, pull data, disconnect a company, etc.). Leave it
         * out for actions that aren't tied to either ERP (login, payment,
         * logout, disconnect-everything).
         * @param {string} message
         * @param {"success"|"error"} type
         * @param {string} [detail] - optional second line shown under the toast title and in the bell drawer
         * @param {"quickbooks"|"xero"} [provider] - which ERP this belongs to, if any
         */
        showStatus(message, type, detail, provider) {
            if (typeof message === "string" && !message.trim().endsWith("...")) {
                NotificationService.add(message, type, detail, provider);
            }
        },

        // Setup/Pull step completion (the green checkmarks on steps 2 and 3)
        // persists per provider + company, same reasoning as the log
        // console: without this, a taskpane refresh would reset every step
        // indicator to "not done" even though the log clearly shows Setup
        // and Pull already ran — the two must stay consistent. Step 1
        // (Connect) isn't stored here at all; it's derived live from
        // AppState.erpConnected, which is already persisted separately.
        STEP_STORAGE_KEY: "fa_step_state",

        /** @returns {Object<string, {setup?:boolean, pull?:boolean}>} */
        _loadStepState() {
            try {
                const raw = localStorage.getItem(this.STEP_STORAGE_KEY);
                const parsed = raw ? JSON.parse(raw) : {};
                return (parsed && typeof parsed === "object" && !Array.isArray(parsed)) ? parsed : {};
            } catch (_) {
                return {};
            }
        },

        _saveStepState(state) {
            try {
                localStorage.setItem(this.STEP_STORAGE_KEY, JSON.stringify(state));
            } catch (_) {
                // Storage full/unavailable — step completion just won't
                // persist across a refresh, but the app keeps working.
            }
        },

        _stepStateKey(provider, companyId) {
            return `${provider}::${companyId || "none"}`;
        },

        /**
         * Applies step-completion state to every step indicator for the
         * current provider + active company — both the "provider-selected"
         * (provStepX) and "connected" (stepX/xeroStepX) sets of elements.
         * Connect is derived live from AppState.erpConnected; Setup/Pull
         * come from persisted per-company storage. Explicitly clears the
         * "complete" class when a step is NOT done for this company, so
         * switching to (or refreshing on) a company that's never had
         * Setup/Pull run never shows a stale green checkmark left over
         * from a previously active company.
         */
        applyStepState() {
            const isQB = AppState.currentProvider === "quickbooks";
            const provider = isQB ? "quickbooks" : "xero";
            const key = this._stepStateKey(provider, AppState.currentCompanyId);
            const state = this._loadStepState()[key] || {};
            const connectDone = !!AppState.erpConnected;

            [
                [isQB ? "stepConnect" : "xeroStepConnect", connectDone],
                [isQB ? "stepSetup"   : "xeroStepSetup",   !!state.setup],
                [isQB ? "stepPull"    : "xeroStepPull",    !!state.pull],
                ["provStepConnect", connectDone],
                ["provStepSetup",   !!state.setup],
                ["provStepPull",    !!state.pull]
            ].forEach(([id, done]) => {
                document.getElementById(id)?.classList.toggle("complete", done);
            });
        },

        /**
         * Marks Setup or Pull complete for the current provider + active
         * company — persists it so it survives a refresh, then re-applies
         * step state to the DOM immediately. "connect" is a no-op here
         * since it's derived live from AppState.erpConnected instead.
         * @param {"connect"|"setup"|"pull"} stepName
         */
        markStepComplete(stepName) {
            if (stepName === "setup" || stepName === "pull") {
                const provider = AppState.currentProvider === "quickbooks" ? "quickbooks" : "xero";
                const key = this._stepStateKey(provider, AppState.currentCompanyId);
                const state = this._loadStepState();
                state[key] = { ...(state[key] || {}), [stepName]: true };
                this._saveStepState(state);
            }
            this.applyStepState();
        },

        /**
         * Marks a progress step as complete (connected-dashboard console).
         * @param {string} step - base step ID ("stepConnect"|"stepSetup"|"stepPull")
         */
        completeStep(step) {
            const stepName = step === "stepSetup" ? "setup" : step === "stepPull" ? "pull" : "connect";
            this.markStepComplete(stepName);
        },

        resetSteps() {
            ["stepConnect","stepSetup","stepPull","xeroStepConnect","xeroStepSetup","xeroStepPull",
             "provStepConnect","provStepSetup","provStepPull"]
                .forEach(id => document.getElementById(id)?.classList.remove("complete", "active"));
        },

        /**
         * Shows the error view with a given message.
         * @param {string} message
         */
        showError(message) {
            const msgEl = document.getElementById("errorMessage");
            if (msgEl) msgEl.textContent = message;
            ViewRouter.show("Error");
            // Also surface as an immediate toast + bell entry, same as any
            // other action failure — the full-screen Error view is shown
            // too, but the user shouldn't have to rely on that alone.
            NotificationService.add(message, "error");
        },

        /**
         * Launches the ERP OAuth popup for the given provider.
         * @param {"quickbooks"|"xero"} provider
         */
        launchERPOAuth(provider) {
            this.showConnecting(provider);
            AppState.currentProvider = provider;
            const isQB  = provider === "quickbooks";
            const pName = isQB ? "QuickBooks" : "Xero";
            const encodedMail = encodeURIComponent(AppState.userEmail || "");
            const tokenParam  = AppState.jwtToken ? `&token=${encodeURIComponent(AppState.jwtToken)}` : "";
            const connectUrl = isQB
                ? `${ApiService.BASE}/api/quickbooks/connect/?tier=${AppState.currentTier}&mail=${encodedMail}${tokenParam}`
                : `${ApiService.BASE}/api/xero/connect?tier=${AppState.currentTier}&mail=${encodedMail}${tokenParam}`;

            // Popup opening is not a completed action — only logged, never toasted.
            this.addLog(`Opening ${pName} sign-in...`);

            // Note: The simulated setTimeout was removed so the UI waits for the real OAuth callback.

            // Guards against the completion/cancellation handling running
            // twice for the same attempt (e.g. the real "connected" message
            // arrives right as the window-closed poll also fires) — only
            // the first one wins, so exactly one outcome is ever processed
            // per attempt.
            let settled = false;
            const finishOnce = (fn) => {
                if (settled) return;
                settled = true;
                fn();
            };

            if (typeof Office !== "undefined" && Office.context && Office.context.ui) {
                Office.context.ui.displayDialogAsync(
                    connectUrl,
                    { height: 60, width: 45, displayInIframe: false, promptBeforeOpen: true },
                    (asyncResult) => {
                        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                            const win = window.open(connectUrl, "_blank", "width=800,height=600");
                            if (!win) {
                                this.showStatus(`Unable to open ${pName} sign-in. Allow popups.`, "error", null, provider);
                            } else {
                                const timer = setInterval(() => {
                                    if (win.closed) {
                                        clearInterval(timer);
                                        // Window closed with no completion message received —
                                        // treat as a plain user cancellation: no verification,
                                        // no callback, no toast. Just restore the prior screen.
                                        finishOnce(() => DashboardService.cancelERPConnection(provider));
                                    }
                                }, 1000);
                            }
                        } else {
                            const dialog = asyncResult.value;
                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                                if (arg.message === "qb_connected" || arg.message === "xero_connected") {
                                    dialog.close();
                                    finishOnce(() => DashboardService.onERPConnected(provider));
                                }
                            });
                            // Fallback: dialog closed manually (error 12006) without a
                            // completion message — this is a user cancellation, not a
                            // failure. No backend verification, no callback, no toast:
                            // just silently restore the screen the user started from.
                            dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
                                if (arg.error === 12006) {
                                    finishOnce(() => DashboardService.cancelERPConnection(provider));
                                }
                            });
                        }
                    }
                );
            } else {
                // Standard popup (browser context)
                const msgHandler = (event) => {
                    if (event.data === "qb_connected" || event.data === "xero_connected") {
                        window.removeEventListener("message", msgHandler);
                        finishOnce(() => DashboardService.onERPConnected(provider));
                    }
                };
                window.addEventListener("message", msgHandler);

                const win = window.open(connectUrl, `${provider}_auth`, "width=800,height=600");
                if (!win) {
                    this.showStatus(`Unable to open ${pName} sign-in. Allow popups.`, "error", null, provider);
                } else {
                    // Fallback: window closed manually without a completion
                    // message — this is a user cancellation. No backend
                    // verification, no completion callback, no toast: just
                    // silently restore the screen the user started from.
                    const timer = setInterval(() => {
                        if (win.closed) {
                            clearInterval(timer);
                            window.removeEventListener("message", msgHandler);
                            finishOnce(() => DashboardService.cancelERPConnection(provider));
                        }
                    }, 1000);
                }
            }
        },

        /**
         * Called when the user manually closes the OAuth popup/dialog
         * without completing authentication. This is a pure cancellation:
         * no backend verification, no onERPConnected callback, and no
         * success/error/warning toast of any kind — closing the popup is
         * not an outcome that gets reported to the user. It restores the
         * dashboard to its real state (the disconnected provider-choice
         * screen, since no connection exists) rather than leaving the
         * "Not connected" provider-selected screen with its Setup/Pull
         * buttons visible — that intermediate screen only makes sense
         * once a connection attempt is in flight, not after it's been
         * abandoned. Only a genuine OAuth callback from the backend
         * (handled by onERPConnected) ever proceeds to verification and
         * notifications.
         * @param {"quickbooks"|"xero"} provider
         */
        cancelERPConnection(provider) {
            void provider; // kept for signature symmetry with the completion path
            this.renderERPSection();
        },

        /**
         * Handles successful ERP OAuth callback — saves state and refreshes dashboard.
         * @param {"quickbooks"|"xero"} provider
         */
        onERPConnected(provider) {
            const isQB = provider === "quickbooks";
            const now  = new Date().toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" });

            AppState.erpConnected     = true;
            AppState.erpType          = provider;
            AppState.erpConnectedDate = now;
            AppState.currentProvider  = provider;

            localStorage.setItem("fa_erp_connected", "true");
            localStorage.setItem("fa_erp_type",      provider);
            localStorage.setItem("fa_erp_date",      now);

            // Attempt to fetch org name from backend. Uses apiFetch (not raw
            // fetch) so the JWT is actually attached — this endpoint is
            // authenticated, so without it every call here 401'd
            // unconditionally and silently fell back to a random ID below.
            const tokenPath = isQB ? "/api/quickbooks/tokens/" : "/api/xero/tokens";

            ApiService.apiFetch(tokenPath)
                .then(r => (r.ok ? r.json() : { tokens: [] }))
                .then(data => {
                    const tokens = data.tokens || [];
                    const realmId = tokens[0]?.realm_id || tokens[0]?.tenant_name || null;
                    // Use backend realm_id if available, else generate a random 16-digit ID
                    const connId = realmId || DashboardService._generateConnectionId();
                    AppState.erpOrgName   = connId;
                    AppState.connectionId = connId;
                    localStorage.setItem("fa_erp_org", connId);
                    DashboardService._finalizeConnection(provider, connId);
                })
                .catch(() => {
                    // Backend not available — generate a random connection ID
                    const connId = DashboardService._generateConnectionId();
                    AppState.erpOrgName   = connId;
                    AppState.connectionId = connId;
                    localStorage.setItem("fa_erp_org", connId);
                    DashboardService._finalizeConnection(provider, connId);
                });
        },

        /**
         * Generates a random 16-digit numeric connection/realm ID.
         * @returns {string}
         */
        _generateConnectionId() {
            // Generate 16-digit number similar to QuickBooks realm IDs
            const part1 = Math.floor(1000000000 + Math.random() * 9000000000); // 10 digits
            const part2 = Math.floor(100000 + Math.random() * 900000);          // 6 digits
            return String(part1) + String(part2);
        },

        /**
         * Finalises connection: marks step 1 complete, renders connected dashboard, updates ID.
         * @param {string} provider
         * @param {string} connId
         */
        _finalizeConnection(provider, connId) {
            const isQB = provider === "quickbooks";

            // Transition to fully connected dashboard (Image 3)
            this.render();

            // Explicitly set the realm ID text in connected header
            const realmEl = document.getElementById("connRealmId");
            if (realmEl) realmEl.textContent = connId;

            // Show success status
            this.showStatus(`${isQB ? "QuickBooks" : "Xero"} connected successfully.`, "success", null, provider);

            // Step 1 (Connect) is derived live from AppState.erpConnected
            // (already true by this point) — this applies it to both the
            // provider-selected and connected step indicators.
            this.applyStepState();
        },

        /**
         * Disconnects the ERP provider — clears state but keeps FinAccrual subscription.
         */
        async disconnectERP() {
            // Optimistically update AppState
            AppState.erpConnected     = false;
            AppState.erpType          = null;
            AppState.erpOrgName       = null;
            AppState.erpConnectedDate = null;
            AppState.isConnected      = false;
            AppState.connectionId     = null;
            AppState.currentCompanyId = null;
            AppState.forceWelcome     = true;

            localStorage.removeItem("fa_erp_connected");
            localStorage.removeItem("fa_erp_type");
            localStorage.removeItem("fa_erp_org");
            localStorage.removeItem("fa_erp_date");

            // resetSteps() clears every step indicator (provider-selected
            // and connected consoles, both providers) — and, since this is
            // an explicit full disconnect (not a refresh), also wipe the
            // persisted Setup/Pull completion state behind them for real.
            this.resetSteps();
            this._saveStepState({});

            // Clear logs — same reasoning: a real disconnect wipes history
            // for good, unlike a taskpane refresh which must preserve it.
            const qbLog = document.getElementById("qbLog");
            const xeroLog = document.getElementById("xeroLog");
            if (qbLog) qbLog.innerHTML = "";
            if (xeroLog) xeroLog.innerHTML = "";
            this._saveLogs([]);

            this.showStatus("Disconnecting all companies...", "success");

            // Disconnect ALL companies for this user from the backend
            try {
                const mail = AppState.userEmail || "";
                const connsRes = await ApiService.apiFetch(`/api/connections?mail=${encodeURIComponent(mail)}`);
                const conns = await connsRes.json();
                const activeConns = (conns || []).filter(c => c.status !== 'Disconnected');
                await Promise.all(activeConns.map(c =>
                    ApiService.apiFetch(`/api/connections/${c.companyId}`, { method: "DELETE" }).catch(() => {})
                ));
            } catch (_) {}

            try { await ExcelService.clearMasterData(); } catch (_) {}

            // Now re-render — all companies should be Disconnected, so it shows the disconnected view
            this.renderERPSection();
            this.showStatus("ERP disconnected. Your FinAccrual account is still active.", "success");
        }
    };

    // ============================================================
    // 8. MAIN APP CONTROLLER — EVENT BINDING & INIT
    // ============================================================
    const AppController = {

        init() {
            // Note: the QuickBooks/Xero log consoles are populated by
            // DashboardService.renderActiveLogConsole(), called once the
            // active provider/company are resolved during restoreSession()
            // below (via render() -> renderERPSection()) — not here, since
            // AppState.currentCompanyId isn't known yet at this point and
            // the console must be scoped to the correct company from the
            // start, never showing another company's history.
            this.bindWelcomeView();
            this.bindPlansView();
            this.bindPaymentView();
            this.bindSuccessView();
            this.bindDashboardView();
            this.bindErrorView();
            this.restoreSession();
        },

        // ---- Welcome View ----
        bindWelcomeView() {
            document.getElementById("btnSignInGoogle")?.addEventListener("click", () => {
                const btn = document.getElementById("btnSignInGoogle");
                if (btn) btn.disabled = true;
                AuthService.openGooglePopup();
                setTimeout(() => { if (btn) btn.disabled = false; }, 3000);
            });

            document.getElementById("btnSignInMicrosoft")?.addEventListener("click", () => {
                const btn = document.getElementById("btnSignInMicrosoft");
                if (btn) btn.disabled = true;
                AuthService.openMicrosoftPopup();
                setTimeout(() => { if (btn) btn.disabled = false; }, 3000);
            });
        },

        // ---- Plans View ----
        bindPlansView() {
            const toggle     = document.getElementById("billingCycleToggle");
            const monthLabel = document.getElementById("labelMonthly");
            const yearLabel  = document.getElementById("labelYearly");

            const updatePrices = (isYearly) => {
                document.querySelectorAll("[data-monthly][data-yearly]").forEach(btn => {
                    const monthly = parseInt(btn.dataset.monthly);
                    const yearly  = parseInt(btn.dataset.yearly);
                    let amountId = "proAmount";
                    if (btn.id === "btnSelectBasic") amountId = "basicAmount";
                    else if (btn.id === "btnSelectStandard") amountId = "standardAmount";
                    
                    const el = document.getElementById(amountId);
                    if (el) el.textContent = isYearly ? yearly : monthly;

                    btn.dataset.activePrice = String(isYearly ? yearly : monthly);
                    btn.dataset.activeCycle = isYearly ? "Yearly" : "Monthly";
                });
                if (monthLabel) monthLabel.classList.toggle("active-label", !isYearly);
                if (yearLabel)  yearLabel.classList.toggle("active-label",  isYearly);
            };

            // Initialize labels
            if (monthLabel) monthLabel.classList.add("active-label");

            if (toggle) {
                toggle.addEventListener("change", () => updatePrices(toggle.checked));
            }

            // Plan select buttons
            document.querySelectorAll("[data-plan]").forEach(btn => {
                btn.addEventListener("click", () => {
                    const plan  = btn.dataset.plan;
                    const cycle = btn.dataset.activeCycle || "Monthly";
                    const price = btn.dataset.activePrice || btn.dataset.monthly || "Custom";

                    if (plan === "Enterprise") {
                        DashboardService.showStatus("Enterprise enquiry sent! Our sales team will contact you.", "success");
                        return;
                    }

                    // Populate payment view
                    const payPlanEl  = document.getElementById("paymentPlanName");
                    const payCycleEl = document.getElementById("paymentBillingCycle");
                    const payTotalEl = document.getElementById("paymentTotal");
                    if (payPlanEl)  payPlanEl.textContent  = plan;
                    if (payCycleEl) payCycleEl.textContent = cycle;
                    if (payTotalEl) payTotalEl.textContent = `₹${price}`;

                    AppState.pendingPlan  = plan;
                    AppState.pendingPrice = price;
                    AppState.pendingCycle = cycle;

                    ViewRouter.show("Payment");
                });
            });

            // Back button
            document.getElementById("btnPlansBack")?.addEventListener("click", () => {
                if (AppState.hasSubscription) {
                    ViewRouter.show("Dashboard");
                } else {
                    ViewRouter.show("Welcome");
                }
            });
        },

        // ---- Payment View ----
        bindPaymentView() {
            document.getElementById("btnPaymentBack")?.addEventListener("click", () => {
                ViewRouter.show("Plans");
            });

            document.getElementById("btnOpenCheckout")?.addEventListener("click", () => {
                CheckoutService.openCheckout(
                    AppState.pendingPlan,
                    AppState.pendingPrice,
                    AppState.pendingCycle
                );
            });

            document.getElementById("btnVerifyPayment")?.addEventListener("click", (e) => {
                e.preventDefault();
                CheckoutService.verifyPayment();
            });
        },

        // ---- Success View ----
        bindSuccessView() {
            document.getElementById("btnGotoDashboard")?.addEventListener("click", () => {
                DashboardService.render();
                ViewRouter.show("Dashboard");
            });
        },

        // ---- Dashboard View ----
        bindDashboardView() {
            // Notification bell / drawer / Clear All
            NotificationService.init();

            // Logouts
            const handleLogout = (e) => {
                if (e) e.preventDefault();
                AuthService.logout();
            };
            document.getElementById("btnBlockLogout")?.addEventListener("click", handleLogout);
            document.getElementById("btnDropdownLogout")?.addEventListener("click", handleLogout);
            document.getElementById("btnChangePlan")?.addEventListener("click", () => {
                ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data: ", err));
                ViewRouter.show("Plans");
            });

            // Copy Subscription ID handler
            const handleCopySubId = () => {
                const subIdText = document.getElementById("dashSubId")?.textContent?.trim() || AppState.subscriptionId;
                if (subIdText) {
                    navigator.clipboard.writeText(subIdText).then(() => {
                        const btn = document.getElementById("btnCopySubId");
                        if (btn) {
                            const originalHTML = btn.innerHTML;
                            btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#16a34a" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>`;
                            setTimeout(() => {
                                btn.innerHTML = originalHTML;
                            }, 1800);
                        }
                    }).catch(err => {
                        console.error("Copy failed: ", err);
                    });
                }
            };
            document.getElementById("btnCopySubId")?.addEventListener("click", handleCopySubId);
            
            // Dropdown Toggle
            const toggleDropdown = (e) => {
                e.stopPropagation();
                const dropdown = document.getElementById("accountMenuDropdown");
                if (dropdown) {
                    dropdown.style.display = dropdown.style.display === "none" ? "flex" : "none";
                }
            };
            document.getElementById("dashHeaderAvatarBtn1")?.addEventListener("click", toggleDropdown);
            document.getElementById("dashHeaderAvatarBtn2")?.addEventListener("click", toggleDropdown);
            document.getElementById("dashHeaderAvatarBtn3")?.addEventListener("click", toggleDropdown);
            document.getElementById("dashHeaderMenuBtn1")?.addEventListener("click", toggleDropdown);

            // Hide dropdown when clicking outside
            document.addEventListener("click", (e) => {
                const dropdown = document.getElementById("accountMenuDropdown");
                if (dropdown && !dropdown.contains(e.target)) {
                    dropdown.style.display = "none";
                }
            });

            // Menu item to show block
            document.querySelectorAll(".dropdown-menu-item").forEach(item => {
                if (item.id === "btnDropdownLogout") return;
                item.addEventListener("click", (e) => {
                    const targetId = e.currentTarget.dataset.target;
                    document.querySelectorAll(".detail-block-card").forEach(c => c.style.display = "none");
                    const targetBlock = document.getElementById(targetId);
                    if (targetBlock) {
                        targetBlock.style.display = "block";
                    }
                    const container = document.getElementById("detailBlocksContainer");
                    if (container) {
                        container.style.display = "flex";
                    }
                    const dropdown = document.getElementById("accountMenuDropdown");
                    if (dropdown) dropdown.style.display = "none";
                });
            });

            // Close block buttons
            document.querySelectorAll(".close-block-btn, .close-card-btn").forEach(btn => {
                btn.addEventListener("click", (e) => {
                    const card = e.target.closest(".detail-block-card");
                    if (card) card.style.display = "none";
                    const container = document.getElementById("detailBlocksContainer");
                    if (container) {
                        const anyVisible = Array.from(container.querySelectorAll(".detail-block-card")).some(c => c.style.display !== "none");
                        if (!anyVisible) {
                            container.style.display = "none";
                        }
                    }
                });
            });

            // Close blocks on backdrop click
            const blocksContainer = document.getElementById("detailBlocksContainer");
            if (blocksContainer) {
                blocksContainer.addEventListener("click", (e) => {
                    if (e.target === blocksContainer) {
                        document.querySelectorAll(".detail-block-card").forEach(c => c.style.display = "none");
                        blocksContainer.style.display = "none";
                    }
                });
            }
            
            // Allow user to cancel connection attempt and go back to disconnected state
            document.getElementById("btnLogoutProvider")?.addEventListener("click", () => {
                DashboardService.renderERPSection();
            });

            // Connect QuickBooks — check if accounts already exist in DB first
            document.getElementById("btnConnectQB")?.addEventListener("click", async () => {
                try {
                    const mail = AppState.userEmail || "";
                    const res = await ApiService.apiFetch(`/api/connections?mail=${encodeURIComponent(mail)}`);
                    const allConns = await res.json();
                    const qbConns = (allConns || []).filter(c => (c.platform || "").toLowerCase() === "quickbooks");
                    if (qbConns.length > 0) {
                        // Existing QB accounts found — reactivate the first active one (or first overall)
                        const toActivate = qbConns.find(c => c.status !== 'Disconnected') || qbConns[0];
                        AppState.currentCompanyId = toActivate.companyId;
                        AppState.currentProvider = "quickbooks";
                        AppState.erpType = "quickbooks";
                        await ApiService.apiFetch(`/api/connections/${toActivate.companyId}/activate`, { method: "POST" });
                        DashboardService.renderERPSection();
                        DashboardService.showStatus(`Resumed QuickBooks session for ${toActivate.companyName}`, "success", null, "quickbooks");
                        return;
                    }
                } catch (_) {}
                // No existing accounts — start OAuth
                DashboardService.showProviderSelected("quickbooks");
                DashboardService.launchERPOAuth("quickbooks");
            });

            // Connect Xero — check if accounts already exist in DB first
            document.getElementById("btnConnectXero")?.addEventListener("click", async () => {
                try {
                    const mail = AppState.userEmail || "";
                    const res = await ApiService.apiFetch(`/api/connections?mail=${encodeURIComponent(mail)}`);
                    const allConns = await res.json();
                    const xeroConns = (allConns || []).filter(c => (c.platform || "").toLowerCase() === "xero");
                    if (xeroConns.length > 0) {
                        // Existing Xero accounts found — reactivate the first one
                        const toActivate = xeroConns.find(c => c.status !== 'Disconnected') || xeroConns[0];
                        AppState.currentCompanyId = toActivate.companyId;
                        AppState.currentProvider = "xero";
                        AppState.erpType = "xero";
                        await ApiService.apiFetch(`/api/connections/${toActivate.companyId}/activate`, { method: "POST" });
                        DashboardService.renderERPSection();
                        DashboardService.showStatus(`Resumed Xero session for ${toActivate.companyName}`, "success", null, "xero");
                        return;
                    }
                } catch (_) {}
                // No existing accounts — start OAuth
                DashboardService.showProviderSelected("xero");
                DashboardService.launchERPOAuth("xero");
            });

            // Connect Provider button in provider-selected state — launches OAuth
            document.getElementById("btnConnectProvider")?.addEventListener("click", () => {
                DashboardService.launchERPOAuth(AppState.currentProvider);
            });

            // Disconnect ERP (overall disconnect)
            document.getElementById("btnDisconnectERP")?.addEventListener("click", async () => {
                await DashboardService.disconnectERP();
            });

            // Sub card Change Plan button
            document.getElementById("btnSubChangePlan")?.addEventListener("click", () => {
                ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data: ", err));
                ViewRouter.show("Plans");
            });

            // Add Company buttons (checks plan limits before starting OAuth)
            // Uses the currently active ERP provider so that:
            //   - Xero dashboard  → opens Xero OAuth
            //   - QuickBooks dashboard → opens QuickBooks OAuth
            const handleAddCompanyClick = () => {
                ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data: ", err));
                const currentPlan = (AppState.subscriptionPlan && AppState.subscriptionPlan !== 'null') ? AppState.subscriptionPlan : "Basic";
                const maxAllowed = currentPlan === "Basic" ? 1 : (currentPlan === "Standard" ? 3 : 10);
                // Derive the provider from the active dashboard, fall back to quickbooks
                const provider = AppState.currentProvider || "quickbooks";
                
                console.log("handleAddCompanyClick: provider = " + provider + ", AppState.currentProvider = " + AppState.currentProvider);

                ApiService.apiFetch("/api/connections?mail=" + encodeURIComponent(AppState.userEmail || ""))
                    .then(r => r.json())
                    .then(conns => {
                        const connsForProvider = (conns || []).filter(c => 
                            (c.platform || "").toLowerCase() === provider.toLowerCase()
                        );
                        if (connsForProvider.length >= maxAllowed) {
                            alert(`Your ${currentPlan} Plan limits active ${provider === 'quickbooks' ? 'QuickBooks' : 'Xero'} connections to ${maxAllowed} companies. Please upgrade your plan to connect more companies.`);
                            ViewRouter.show("Plans");
                        } else {
                            DashboardService.showProviderSelected(provider);
                            DashboardService.launchERPOAuth(provider);
                        }
                    })
                    .catch(() => {
                        DashboardService.launchERPOAuth(provider);
                    });
            };
            document.getElementById("btnAddCompany")?.addEventListener("click", handleAddCompanyClick);
            document.getElementById("btnModalAddCompany")?.addEventListener("click", () => {
                const modal = document.getElementById("changeCompanyModal");
                if (modal) modal.style.display = "none";
                handleAddCompanyClick();
            });

            // Change Company Modal open/close & switch logic
            let selectedModalCompanyId = null;
            const openChangeCompanyModal = () => {
                const modal = document.getElementById("changeCompanyModal");
                const modalList = document.getElementById("modalCompanyList");
                if (!modal || !modalList) return;

                ApiService.apiFetch("/api/connections?mail=" + encodeURIComponent(AppState.userEmail || ""))
                    .then(r => r.json())
                    .then(conns => {
                        modalList.innerHTML = "";
                        // Filter to show only companies for the current platform
                        const currentPlatform = AppState.currentProvider || "quickbooks";
                        const platformConns = conns.filter(c => (c.platform || "quickbooks").toLowerCase() === currentPlatform);
                        const platformLabel = currentPlatform === "xero" ? "Xero" : "QuickBooks";

                        // Update modal title
                        const modalTitle = document.querySelector("#changeCompanyModal .fa-modal-title, #changeCompanyModal h3");
                        if (modalTitle) modalTitle.textContent = `Switch ${platformLabel} Company`;

                        platformConns.forEach(c => {
                            const isSelected = c.companyId === (selectedModalCompanyId || AppState.currentCompanyId);
                            const isXero = (c.platform || "").toLowerCase() === "xero";
                            const isDisconnected = c.status === 'Disconnected';
                            const row = document.createElement("div");
                            row.className = `fa-modal-company-row ${isSelected ? "selected" : ""} ${isDisconnected ? "disconnected" : ""}`;
                            row.dataset.companyId = c.companyId;
                            row.dataset.platform = (c.platform || "quickbooks").toLowerCase();
                            
                            const displayName = c.companyName || (isXero ? "Xero Organisation" : "QuickBooks Company");

                            row.innerHTML = `
                                <div class="fa-company-icon ${isXero ? 'xero-company-icon' : ''}">${isXero ? 'xero' : 'qb'}</div>
                                <div class="fa-company-info">
                                    <div class="fa-company-name">${displayName}${isDisconnected ? ' <span style="color:#ef4444;font-size:10px">(Disconnected)</span>' : ''}</div>
                                    <div class="fa-company-tag">${isXero ? "Tenant ID" : "Realm ID"}: ${c.companyId || "—"}</div>
                                </div>
                            `;
                            row.addEventListener("click", () => {
                                modalList.querySelectorAll(".fa-modal-company-row").forEach(r => r.classList.remove("selected"));
                                row.classList.add("selected");
                                selectedModalCompanyId = c.companyId;
                            });
                            modalList.appendChild(row);
                        });
                        modal.style.display = "flex";
                    });
            };

            document.getElementById("btnChangeCompany")?.addEventListener("click", openChangeCompanyModal);
            document.getElementById("btnManageCompanies")?.addEventListener("click", openChangeCompanyModal);

            const closeChangeCompanyModal = () => {
                const modal = document.getElementById("changeCompanyModal");
                if (modal) modal.style.display = "none";
            };
            document.getElementById("btnCloseChangeCompany")?.addEventListener("click", closeChangeCompanyModal);
            document.getElementById("btnCancelChangeCompany")?.addEventListener("click", closeChangeCompanyModal);

            document.getElementById("btnConfirmChangeCompany")?.addEventListener("click", async () => {
                if (selectedModalCompanyId) {
                    AppState.currentCompanyId = selectedModalCompanyId;
                    // Determine the platform from the selected modal row
                    const selectedRow = document.querySelector(`.fa-modal-company-row[data-company-id="${selectedModalCompanyId}"]`);
                    if (selectedRow && selectedRow.dataset.platform) {
                        AppState.currentProvider = selectedRow.dataset.platform;
                        AppState.erpType = selectedRow.dataset.platform;
                    }
                    const switchedPlatform = AppState.currentProvider;
                    ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data: ", err));
                    try {
                        await ApiService.apiFetch(`/api/connections/${selectedModalCompanyId}/activate`, { method: "POST" });
                    } catch (_) {}
                    DashboardService.renderERPSection();
                    DashboardService.showStatus("Active company switched successfully.", "success", null, switchedPlatform);
                }
                closeChangeCompanyModal();
            });

            // Modal search filter
            document.getElementById("companySearchInput")?.addEventListener("input", (e) => {
                const term = e.target.value.toLowerCase();
                document.querySelectorAll(".fa-modal-company-row").forEach(row => {
                    const text = row.textContent.toLowerCase();
                    row.style.display = text.includes(term) ? "flex" : "none";
                });
            });

            // Close context menu when clicking anywhere else
            document.addEventListener("click", (e) => {
                const menu = document.getElementById("companyContextMenu");
                if (menu && !menu.contains(e.target) && !e.target.classList.contains("fa-btn-dots")) {
                    menu.style.display = "none";
                }
            });

            // Disconnect Active Company
            document.getElementById("btnDisconnectActiveCompany")?.addEventListener("click", async () => {
                const companyId = AppState.currentCompanyId;
                if (!companyId) return;
                const disconnectPlatform = AppState.currentProvider;

                DashboardService.showStatus("Disconnecting company...", "success", null, disconnectPlatform);
                try {
                    const res = await ApiService.apiFetch(`/api/connections/${companyId}`, {
                        method: "DELETE"
                    });
                    if (res.ok) {
                        AppState.currentCompanyId = null; // Reset so a new active company gets picked
                        DashboardService.showStatus("Company disconnected successfully.", "success", null, disconnectPlatform);
                        DashboardService.renderERPSection();
                    } else {
                        DashboardService.showStatus("Failed to disconnect company.", "error", null, disconnectPlatform);
                    }
                } catch (err) {
                    DashboardService.showStatus("Error disconnecting company.", "error", null, disconnectPlatform);
                }
            });

            // Dropdown selection change
            document.getElementById("companySelectDropdown")?.addEventListener("change", (e) => {
                const dropdown = e.target;
                const opt = dropdown.options[dropdown.selectedIndex];
                if (opt) {
                    AppState.currentCompanyId = opt.value;
                    AppState.currentProvider = opt.dataset.platform;
                    AppState.erpType = opt.dataset.platform;
                    ExcelService.clearMasterData().catch(err => console.error("Error clearing Excel data: ", err));

                    // Activate in backend and re-render
                    ApiService.apiFetch(`/api/connections/${opt.value}/activate`, { method: "POST" })
                        .then(() => DashboardService.renderERPSection())
                        .catch(() => DashboardService.renderERPSection());

                    DashboardService.addLog(`Switched active company to: ${opt.textContent}`);
                }
            });

            // Setup Sheets button in provider-selected state
            document.getElementById("setupBtnProv")?.addEventListener("click", async () => {
                try {
                    document.getElementById("provStepSetup")?.classList.add("active");
                    DashboardService.addLog(`Setting up Master & Input sheets for ${AppState.currentProvider === "quickbooks" ? "QuickBooks" : "Xero"}...`);
                    DashboardService.showStatus("Setting up sheets...", "success", null, AppState.currentProvider);
                    await ExcelService.setupWorkbookSheets(AppState.currentProvider);
                    DashboardService.markStepComplete("setup");
                    DashboardService.addLog("Sheets setup successfully.");
                    DashboardService.showStatus("Master and Input sheets setup successfully.", "success", null, AppState.currentProvider);
                } catch (error) {
                    console.error(error);
                    DashboardService.addLog("Error setting up sheets: " + error.message);
                    DashboardService.showStatus("Error setting up sheets.", "error", null, AppState.currentProvider);
                }
            });

            // Pull Data button in provider-selected state
            document.getElementById("pullBtnProv")?.addEventListener("click", async () => {
                try {
                    document.getElementById("provStepPull")?.classList.add("active");
                    DashboardService.addLog(`Pulling master data from ${AppState.currentProvider === "quickbooks" ? "QuickBooks" : "Xero"}...`);
                    DashboardService.showStatus("Pulling data...", "success", null, AppState.currentProvider);
                    const data = await ApiService.fetchMasterData(AppState.currentProvider, AppState.currentCompanyId);
                    await ExcelService.writeMasterData(AppState.currentProvider, data);
                    DashboardService.markStepComplete("pull");
                    const changedCount = countChangedMasterDataRecords(data);
                    const pullTitle = "Data pull completed successfully.";
                    const pullDetail = changedCount > 0
                        ? `${changedCount} new/updated record${changedCount === 1 ? "" : "s"} found.`
                        : "Data synchronized successfully.";
                    DashboardService.addLog(`${pullTitle} ${pullDetail}`);
                    DashboardService.showStatus(pullTitle, "success", pullDetail, AppState.currentProvider);
                    DashboardService.renderERPSection();
                } catch (error) {
                    console.error(error);
                    // ApiService already showed the orange "reconnect" banner
                    // (or the offline banner) as a global side effect when
                    // this came from apiFetch — branch on the standardized
                    // `.code` here too, never on message text.
                    const isExpired = error.code === ERROR_CODES.ERP_SESSION_EXPIRED;
                    const msg = isExpired
                        ? error.message
                        : "Error pulling data: " + error.message;
                    DashboardService.addLog(msg);
                    DashboardService.showStatus(
                        isExpired ? error.message : "Data pull failed.",
                        "error",
                        isExpired ? "" : "Please try again.",
                        AppState.currentProvider
                    );
                    if (isExpired) {
                        // renderERPConsole() only toggles a progress-step
                        // marker — it doesn't touch the company badge. To
                        // actually flip ACTIVE -> Reconnect in the UI, we
                        // need to re-fetch connections and re-render the
                        // company list, which is what renderERPSection()
                        // does. Guarded in case of an unexpected error, so
                        // it can't surface as an uncaught runtime popup.
                        try {
                            DashboardService.renderERPSection();
                        } catch (renderErr) {
                            console.error("Failed to refresh ERP section after session expiry:", renderErr);
                        }
                    }
                }
            });

            // Journal type buttons in provider-selected state
            document.querySelectorAll("#dashProviderSelected .conn-journal-btn").forEach(btn => {
                btn.addEventListener("click", () => {
                    document.querySelectorAll("#dashProviderSelected .conn-journal-btn").forEach(b => b.classList.remove("active"));
                    btn.classList.add("active");
                });
            });

            // Setup Sheets button
            document.getElementById("setupBtn")?.addEventListener("click", async () => {
                try {
                    const stepSetupId = AppState.currentProvider === "quickbooks" ? "stepSetup" : "xeroStepSetup";
                    document.getElementById(stepSetupId)?.classList.add("active");

                    DashboardService.addLog(`Setting up Master & Input sheets for ${AppState.currentProvider === "quickbooks" ? "QuickBooks" : "Xero"}...`);
                    DashboardService.showStatus("Setting up sheets...", "success", null, AppState.currentProvider);
                    await ExcelService.setupWorkbookSheets(AppState.currentProvider);
                    DashboardService.completeStep("stepSetup");
                    DashboardService.addLog("Sheets setup successfully.");
                    DashboardService.showStatus("Master and Input sheets setup successfully.", "success", null, AppState.currentProvider);
                } catch (error) {
                    console.error(error);
                    DashboardService.addLog("Error setting up sheets: " + error.message);
                    DashboardService.showStatus("Error setting up sheets.", "error", null, AppState.currentProvider);
                }
            });

            // Pull Data button
            document.getElementById("pullBtn")?.addEventListener("click", async () => {
                try {
                    const stepPullId = AppState.currentProvider === "quickbooks" ? "stepPull" : "xeroStepPull";
                    document.getElementById(stepPullId)?.classList.add("active");

                    DashboardService.addLog(`Pulling master data from ${AppState.currentProvider === "quickbooks" ? "QuickBooks" : "Xero"}...`);
                    DashboardService.showStatus("Pulling data...", "success", null, AppState.currentProvider);
                    const data = await ApiService.fetchMasterData(AppState.currentProvider, AppState.currentCompanyId);
                    await ExcelService.writeMasterData(AppState.currentProvider, data);
                    DashboardService.completeStep("stepPull");
                    const changedCount = countChangedMasterDataRecords(data);
                    const pullTitle = "Data pull completed successfully.";
                    const pullDetail = changedCount > 0
                        ? `${changedCount} new/updated record${changedCount === 1 ? "" : "s"} found.`
                        : "Data synchronized successfully.";
                    DashboardService.addLog(`${pullTitle} ${pullDetail}`);
                    DashboardService.showStatus(pullTitle, "success", pullDetail, AppState.currentProvider);
                    DashboardService.renderERPSection();
                } catch (error) {
                    console.error(error);
                    // ApiService already showed the orange "reconnect" banner
                    // (or the offline banner) as a global side effect when
                    // this came from apiFetch — branch on the standardized
                    // `.code` here too, never on message text.
                    const isExpired = error.code === ERROR_CODES.ERP_SESSION_EXPIRED;
                    const msg = isExpired
                        ? error.message
                        : "Error pulling data: " + error.message;
                    DashboardService.addLog(msg);
                    DashboardService.showStatus(
                        isExpired ? error.message : "Please set up the master sheet before pulling the master data",
                        "error",
                        isExpired ? "" : "",
                        AppState.currentProvider
                    );
                    if (isExpired) {
                        // renderERPConsole() only toggles a progress-step
                        // marker — it doesn't touch the company badge. To
                        // actually flip ACTIVE -> Reconnect in the UI, we
                        // need to re-fetch connections and re-render the
                        // company list, which is what renderERPSection()
                        // does. Guarded in case of an unexpected error, so
                        // it can't surface as an uncaught runtime popup.
                        try {
                            DashboardService.renderERPSection();
                        } catch (renderErr) {
                            console.error("Failed to refresh ERP section after session expiry:", renderErr);
                        }
                    }
                }
            });

            // Journal type buttons
            document.querySelectorAll(".conn-journal-btn").forEach(btn => {
                btn.addEventListener("click", () => {
                    document.querySelectorAll(".conn-journal-btn").forEach(b => b.classList.remove("active"));
                    btn.classList.add("active");
                });
            });

            // Provider Tab Toggles
            const tabQB   = document.getElementById("tabQB");
            const tabXero = document.getElementById("tabXero");
            const options = document.getElementById("connTierOptions");

            const toggleOptions = (provider) => {
                if (options) {
                    options.style.display = options.style.display === "flex" ? "none" : "flex";
                    options.style.flexDirection = "column";
                    
                    // Render correct options list
                    const isQB = provider === "quickbooks";
                    document.getElementById("tierBasic").textContent = isQB ? "Q Basic" : "Xero Basic";
                    document.getElementById("tierStandard").textContent = isQB ? "Q Standard" : "Xero Standard";
                    document.getElementById("tierPro").textContent = isQB ? "Q Pro" : "Xero Pro";
                }
            };

            tabQB?.addEventListener("click", (e) => {
                e.stopPropagation();
                if (AppState.erpConnected && AppState.erpType !== "quickbooks") return; // Tab locked to connection
                toggleOptions("quickbooks");
            });

            tabXero?.addEventListener("click", (e) => {
                e.stopPropagation();
                if (AppState.erpConnected && AppState.erpType !== "xero") return; // Tab locked to connection
                toggleOptions("xero");
            });

            // Dropdown option clicks
            document.querySelectorAll(".conn-tier-opt").forEach(opt => {
                opt.addEventListener("click", (e) => {
                    e.stopPropagation();
                    document.querySelectorAll(".conn-tier-opt").forEach(o => o.classList.remove("selected"));
                    opt.classList.add("selected");
                    
                    const tierName = opt.textContent;
                    AppState.erpTier = tierName;
                    const badge = document.getElementById("connTierBadge");
                    if (badge) badge.textContent = tierName;

                    if (options) options.style.display = "none";
                });
            });

            // Collapse dropdown on outside click
            document.addEventListener("click", () => {
                if (options) options.style.display = "none";
            });

            // Refresh Schedule buttons
            const handleRefreshClick = async (event) => {
                const button = event.currentTarget;
                const isProv = button.id === "btnRefreshScheduleProv";
                
                // Verify that both setup and pull steps are complete first
                let isSetupComplete = false;
                let isPullComplete = false;
                
                if (isProv) {
                    isSetupComplete = document.getElementById("provStepSetup")?.classList.contains("complete");
                    isPullComplete = document.getElementById("provStepPull")?.classList.contains("complete");
                } else {
                    const stepSetupId = AppState.currentProvider === "quickbooks" ? "stepSetup" : "xeroStepSetup";
                    const stepPullId = AppState.currentProvider === "quickbooks" ? "stepPull" : "xeroStepPull";
                    isSetupComplete = document.getElementById(stepSetupId)?.classList.contains("complete");
                    isPullComplete = document.getElementById(stepPullId)?.classList.contains("complete");
                }

                if (!isSetupComplete || !isPullComplete) {
                    DashboardService.addLog("Cannot refresh: Setup and initial Pull must be completed first.");
                    DashboardService.showStatus("Please pull the master data before refreshing the schedule.", "error", null, AppState.currentProvider);
                    return;
                }

                const icon = button.querySelector(".refresh-icon");
                if (icon) icon.classList.add("spin");

                try {
                    DashboardService.addLog(`Refreshing live data from ${AppState.currentProvider === "quickbooks" ? "QuickBooks" : "Xero"}...`);
                    DashboardService.showStatus("Refreshing...", "success", null, AppState.currentProvider);

                    const data = await ApiService.fetchMasterData(AppState.currentProvider, AppState.currentCompanyId);
                    await ExcelService.writeMasterData(AppState.currentProvider, data);

                    const timestamp = new Date().toLocaleTimeString();
                    await ExcelService.stampLastRefreshed(timestamp);

                    const changedCount = countChangedMasterDataRecords(data);
                    const refreshTitle = "Data refresh completed successfully.";
                    const refreshDetail = changedCount > 0
                        ? `${changedCount} new/updated record${changedCount === 1 ? "" : "s"} found.`
                        : "Data synchronized successfully.";
                    DashboardService.addLog("Live data refreshed and timestamp stamped successfully.");
                    DashboardService.showStatus(refreshTitle, "success", refreshDetail, AppState.currentProvider);
                } catch (err) {
                    console.error("Refresh error:", err);
                    DashboardService.addLog("Error refreshing: " + err.message);
                    DashboardService.showStatus("Data refresh failed.", "error", "Please try again.", AppState.currentProvider);
                } finally {
                    setTimeout(() => {
                        if (icon) icon.classList.remove("spin");
                    }, 1000);
                }
            };

            document.getElementById("btnRefreshScheduleProv")?.addEventListener("click", handleRefreshClick);
            document.getElementById("btnRefreshSchedule")?.addEventListener("click", handleRefreshClick);
        },

        // ---- Error View ----
        bindErrorView() {
            document.getElementById("btnRetry")?.addEventListener("click", () => {
                ViewRouter.show("Welcome");
            });
        },

        // ---- Session Restoration ----
        /**
         * On load, checks localStorage to see if the user already has an active session.
         * If so, skips the welcome screen and navigates straight to the dashboard.
         */
        restoreSession() {
            const email = localStorage.getItem("fa_user_email");
            if (email) {
                AppState.userEmail = email;
                AppState.userName = localStorage.getItem("fa_user_name");
                AppState.userProvider = localStorage.getItem("fa_user_provider");
                AppState.hasSubscription = localStorage.getItem("fa_has_subscription") === "true";
                AppState.subscriptionId = localStorage.getItem("fa_subscription_id");
                AppState.subscriptionPlan = (v => (!v || v === 'null' || v === 'undefined') ? null : v)(localStorage.getItem("fa_subscription_plan"));
                AppState.erpConnected = localStorage.getItem("fa_erp_connected") === "true";
                AppState.erpType = localStorage.getItem("fa_erp_type");
                AppState.currentCompanyId = localStorage.getItem("fa_current_company_id") || null;

                DashboardService.render();

                const lastView = localStorage.getItem("fa_last_view");
                if (lastView && lastView !== "Welcome" && lastView !== "Loading" && lastView !== "Error") {
                    ViewRouter.show(lastView);
                } else {
                    ViewRouter.show("Dashboard");
                }

                // Cached localStorage can go stale (plan changed from
                // another device/session), so use /api/auth/me to pick up
                // a newer confirmed plan. This must only ever CONFIRM or
                // UPGRADE the displayed plan — it must never clear it back
                // to "no plan" (which renders as Basic), since a failed
                // request, an offline moment, or a write that simply
                // hasn't landed yet would otherwise look exactly like the
                // plan silently reverting to Basic on refresh.
                if (AppState.jwtToken) {
                    ApiService.apiFetch("/api/auth/me")
                        .then(r => (r.ok ? r.json() : null))
                        .then(result => {
                            const dbPlan = result?.user?.plan || null;
                            if (dbPlan && dbPlan !== AppState.subscriptionPlan) {
                                AppState.subscriptionPlan = dbPlan;
                                AppState.hasSubscription = true;
                                localStorage.setItem("fa_subscription_plan", dbPlan);
                                localStorage.setItem("fa_has_subscription", "true");
                                DashboardService.render();
                                DashboardService.renderERPSection();
                            }
                            // If dbPlan is empty/missing, do nothing — keep
                            // showing whatever plan was already cached.
                        })
                        .catch(() => {
                            // Offline or request failed — keep showing the
                            // cached plan rather than blocking the UI.
                        });
                }
            } else {
                // Reset AppState to defaults
                AppState.userEmail        = null;
                AppState.userName         = null;
                AppState.userProvider     = null;
                AppState.hasSubscription  = false;
                AppState.subscriptionId   = null;
                AppState.subscriptionPlan = null;
                AppState.erpConnected     = false;
                AppState.erpType          = null;
                AppState.erpOrgName       = null;
                AppState.erpConnectedDate = null;

                // Always show welcome screen
                ViewRouter.show("Welcome");
            }
        }
    };

    // ============================================================
    // 9. BOOT
    // ============================================================
    AppController.init();

});