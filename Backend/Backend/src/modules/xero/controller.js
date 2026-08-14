'use strict';

const querystring = require('querystring');
const config      = require('../../core/config');
const CONSTANTS   = require('../../core/constants');
const logger      = require('../../core/logger');
const { generateOAuthState } = require('../../core/helpers');
const XeroService         = require('./service');
const XeroTokenRepository = require('./repository');
const { ValidationError } = require('../../core/errors/AppError');

/**
 * XeroController
 * -----------------------------------------------------------------
 * Handles all incoming HTTP requests for the Xero module.
 * Delegates all business logic to XeroService.
 * Does NOT contain any data-transformation or mapping logic —
 * that responsibility lives in mapper.js (used by the Service).
 * -----------------------------------------------------------------
 */
class XeroController {

    /**
     * GET /api/xero/connect
     * Generates the Xero OAuth authorization URL and redirects.
     *
     * Requires authentication (the frontend passes the JWT as ?token=...
     * since this is a browser navigation, not a fetch). The owning email
     * is taken exclusively from the verified token (req.user.email), never
     * from a client-suppliable query param — otherwise anyone could
     * initiate a connect flow tagged with someone else's email and inject
     * a connection into that other user's account.
     */
    connectXero = async (req, res, next) => {
        try {
            const { XeroToken } = require('../../core/database');
            const mail = req.user.email;
            const { Op } = require('sequelize');
            const whereClause = { status: { [Op.ne]: 'Disconnected' }, mail };
            const xeroCount = await XeroToken.count({ where: whereClause });
            const tier = (req.query.tier || 'pro').toLowerCase();

            let maxAllowed = 10;
            if (tier === 'trial') maxAllowed = 1;
            else if (tier === 'basic') maxAllowed = 1;
            else if (tier === 'standard') maxAllowed = 3;

            if (xeroCount >= maxAllowed) {
                return res.send(`
                    <html>
                        <body style="font-family:sans-serif; text-align:center; padding: 40px; background:#fff1f2; color:#9f1239;">
                            <div style="font-size: 50px; margin-bottom: 20px;">⚠️</div>
                            <h2>Connection Limit Reached</h2>
                            <p style="font-size: 14px; color: #4b5563;">Your subscription tier (${tier.toUpperCase()}) allows a maximum of ${maxAllowed} connected company.</p>
                            <p style="font-size: 14px; color: #4b5563;">Please disconnect an existing company or upgrade your plan to connect more.</p>
                            <button onclick="window.close()" style="margin-top: 20px; padding:10px 20px; background:#be123c; color:white; border:none; border-radius:5px; cursor:pointer; font-weight: bold;">Close Window</button>
                        </body>
                    </html>
                `);
            }

            const state = generateOAuthState();
            req.session.xero_state = state;
            req.session.user_mail  = mail;
            req.session.xero_tier  = tier;
            req.session.xero_max_allowed = maxAllowed;

            const params = {
                response_type: 'code',
                client_id:     config.XERO.CLIENT_ID,
                redirect_uri:  config.XERO.REDIRECT_URI,
                scope:         config.XERO.SCOPES,
                state
            };

            const authUrl = `${CONSTANTS.XERO.AUTH_URL}?${querystring.stringify(params)}`;
            logger.info('Redirecting to Xero OAuth... URL: ' + authUrl);
            res.redirect(authUrl);
        } catch (error) {
            next(error);
        }
    };

    /**
     * GET /api/xero/callback
     * Exchanges the OAuth code for tokens, fetches ALL available Xero orgs,
     * stores tokens temporarily in the session, then returns a company-selection
     * HTML page so the user can choose which orgs to activate.
     */
    xeroCallback = async (req, res, next) => {
        try {
            const { code } = req.query;
            const mail        = req.session?.user_mail || req.session?.admin?.email || req.session?.googleUser?.email || null;
            const tier        = req.session?.xero_tier || 'pro';
            const maxAllowed  = req.session?.xero_max_allowed || 10;

            // Exchange code → tokens, fetch all orgs (no DB writes yet)
            const { tokens, tenants } = await XeroService.exchangeTokensOnly(code);

            if (tenants.length === 0) {
                return res.status(400).send(`
                    <html><body style="font-family:sans-serif;text-align:center;padding:40px;">
                        <h2>No Xero organisations found.</h2>
                        <p>Make sure your Xero account has at least one organisation.</p>
                        <button onclick="window.close()">Close</button>
                    </body></html>
                `);
            }

            // Store tokens + tenants in session temporarily — we'll use them in selectCompanies
            req.session.xero_pending_tokens  = tokens;
            req.session.xero_pending_tenants = tenants;
            req.session.xero_pending_mail    = mail;

            // Fetch existing connected tokens (Active or Not Synced — i.e.
            // anything short of Disconnected) to pre-check already
            // connected companies on the selection screen.
            const { XeroToken } = require('../../core/database');
            const { Op } = require('sequelize');
            const whereClause = mail ? { mail } : {};
            const existingActive = await XeroToken.findAll({
                where: { ...whereClause, status: { [Op.ne]: 'Disconnected' } }
            });
            const activeTenantIds = new Set(existingActive.map(t => t.tenant_id));

            // Build the tier label
            const tierLabel = tier === 'trial' ? 'Trial (1)' : tier === 'basic' ? 'Basic (1)' : tier === 'standard' ? 'Standard (3)' : 'Pro (10)';

            // Build the company-selection HTML page
            const companyRows = tenants.map(t => {
                const isSelected = activeTenantIds.has(t.tenantId);
                return `
                <div class="company-row ${isSelected ? 'selected already-connected' : ''}" id="row_${t.tenantId}" data-id="${t.tenantId}">
                    <input
                        type="checkbox"
                        class="company-cb"
                        id="cb_${t.tenantId}"
                        value="${t.tenantId}"
                        data-name="${(t.tenantName || 'Xero Organisation').replace(/"/g, '&quot;')}"
                        ${isSelected ? 'checked disabled' : ''}
                    />
                    <div class="company-icon">xero</div>
                    <div class="company-info">
                        <div class="company-name">${t.tenantName || 'Xero Organisation'}${isSelected ? ' <span style="font-size:10px;color:#059669;font-weight:600;">(Already Connected)</span>' : ''}</div>
                        <div class="company-id">Realm: ${t.tenantId}</div>
                    </div>
                </div>
            `;
            }).join('');


            return res.send(`<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Select Xero Companies</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif;
    background: #f4f6f9;
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #1a2035;
  }
  .card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 16px;
    padding: 36px 32px 28px;
    width: 100%;
    max-width: 480px;
    box-shadow: 0 20px 40px rgba(0,0,0,0.08);
  }
  .xero-logo {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 20px;
  }
  .xero-logo-icon {
    width: 38px; height: 38px;
    background: linear-gradient(135deg, #13b5ea, #0d7db0);
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-weight: 700; font-size: 14px; color: #fff; letter-spacing: -0.5px;
  }
  h1 {
    font-size: 20px;
    font-weight: 700;
    color: #1a2035;
    line-height: 1.3;
  }
  .subtitle {
    font-size: 13px;
    color: #64748b;
    margin-top: 6px;
    line-height: 1.5;
  }
  .plan-badge {
    display: inline-block;
    background: rgba(19, 181, 234, 0.15);
    color: #13b5ea;
    border: 1px solid rgba(19, 181, 234, 0.3);
    border-radius: 20px;
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 600;
    margin-top: 10px;
  }
  .company-list {
    margin-top: 20px;
    display: flex;
    flex-direction: column;
    gap: 8px;
    max-height: 300px;
    overflow-y: auto;
    padding-right: 4px;
  }
  .company-list::-webkit-scrollbar { width: 4px; }
  .company-list::-webkit-scrollbar-track { background: transparent; }
  .company-list::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
  .company-row {
    display: flex;
    align-items: center;
    gap: 12px;
    background: #f8fafc;
    border: 1.5px solid #e2e8f0;
    border-radius: 10px;
    padding: 12px 14px;
    cursor: pointer;
    transition: border-color 0.2s, background 0.2s;
    user-select: none;
  }
  .company-row:hover { border-color: #13b5ea; background: #f0f9ff; }
  /* Native checkbox — styled to be visible and large enough to tap */
  .company-cb {
    width: 18px;
    height: 18px;
    accent-color: #13b5ea;
    cursor: pointer;
    flex-shrink: 0;
    margin: 0;
    pointer-events: none; /* Row click handler manages toggle; this prevents double-fire */
  }
  .company-row.selected {
    border-color: #13b5ea;
    background: rgba(19, 181, 234, 0.08);
  }
  .company-row.already-connected {
    opacity: 0.75;
    cursor: not-allowed;
    background: #f1f5f9;
  }
  .company-row.already-connected:hover {
    border-color: #e2e8f0;
    background: #f1f5f9;
  }
  .company-row.selected .company-icon { background: linear-gradient(135deg, #13b5ea, #0d7db0); }
  .company-icon {
    width: 36px; height: 36px;
    background: #e2e8f0;
    border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    font-size: 10px; font-weight: 700;
    color: #64748b;
    flex-shrink: 0;
    transition: background 0.2s;
  }
  .company-row.selected .company-icon { color: #fff; }
  .company-info { flex: 1; min-width: 0; }
  .company-name {
    font-size: 14px;
    font-weight: 600;
    color: #1a2035;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .company-id {
    font-size: 11px;
    color: #64748b;
    margin-top: 2px;
    font-family: 'SFMono-Regular', Consolas, monospace;
  }
  /* cb-visual: shown at the right side as a custom dot indicator */
  .cb-visual {
    width: 22px; height: 22px;
    border: 2px solid #cbd5e1;
    border-radius: 50%;
    flex-shrink: 0;
    display: flex; align-items: center; justify-content: center;
    transition: all 0.2s;
    cursor: pointer;
  }
  .company-row.selected .cb-visual {
    background: #13b5ea;
    border-color: #13b5ea;
  }
  .company-row.selected .cb-visual::after {
    content: '';
    width: 8px; height: 8px;
    background: #fff;
    border-radius: 50%;
    display: block;
  }
  .selection-info {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-top: 14px;
    padding: 8px 12px;
    background: #f8fafc;
    border-radius: 8px;
    border: 1px solid #e2e8f0;
  }
  .selection-count {
    font-size: 12px;
    color: #64748b;
  }
  .selection-count span {
    color: #13b5ea;
    font-weight: 600;
  }
  .limit-warning {
    font-size: 11px;
    color: #d29922;
  }
  .btn-confirm {
    width: 100%;
    margin-top: 18px;
    padding: 13px;
    background: linear-gradient(135deg, #13b5ea, #0d7db0);
    color: #fff;
    border: none;
    border-radius: 10px;
    font-size: 14px;
    font-weight: 600;
    cursor: pointer;
    transition: opacity 0.2s, transform 0.1s;
    letter-spacing: 0.3px;
  }
  .btn-confirm:hover:not(:disabled) { opacity: 0.9; transform: translateY(-1px); }
  .btn-confirm:disabled { opacity: 0.45; cursor: not-allowed; transform: none; }
  .btn-confirm.loading { pointer-events: none; }
  .spinner {
    display: inline-block;
    width: 14px; height: 14px;
    border: 2px solid rgba(255,255,255,0.4);
    border-top-color: #fff;
    border-radius: 50%;
    animation: spin 0.7s linear infinite;
    margin-right: 8px;
    vertical-align: middle;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
  .error-banner {
    display: none;
    background: rgba(248, 81, 73, 0.1);
    border: 1px solid rgba(248, 81, 73, 0.3);
    border-radius: 8px;
    padding: 10px 14px;
    margin-top: 12px;
    font-size: 13px;
    color: #f85149;
  }
</style>
</head>
<body>
<div class="card">
  <div class="xero-logo">
    <div class="xero-logo-icon">xero</div>
    <div>
      <div style="font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:1px;">Connected to</div>
      <div style="font-size:15px;font-weight:700;color:#1a2035;">Xero</div>
    </div>
  </div>

  <h1>Select Your Companies</h1>
  <p class="subtitle">
    Choose the Xero organisations you want to use in FinAccrual.
    Your plan allows up to <strong style="color:#1a2035;">${maxAllowed}</strong> connected ${maxAllowed === 1 ? 'company' : 'companies'}.
  </p>
  <span class="plan-badge">${tierLabel.toUpperCase()} PLAN</span>

  <div class="company-list" id="companyList">
    ${companyRows}
  </div>

  <div class="selection-info">
    <span class="selection-count">Selected: <span id="selCount">0</span> / ${maxAllowed}</span>
    <span class="limit-warning" id="limitWarning" style="display:none;">Limit reached</span>
  </div>

  <div class="error-banner" id="errorBanner"></div>

  <button class="btn-confirm" id="btnConfirm" disabled>
    Connect Selected Companies
  </button>
</div>

<script>
  const MAX = ${maxAllowed};
  let selected = new Set();

  function updateUI() {
    const countEl = document.getElementById('selCount');
    const btn     = document.getElementById('btnConfirm');
    const warning = document.getElementById('limitWarning');
    countEl.textContent = selected.size;
    btn.disabled = selected.size === 0;
    warning.style.display = selected.size >= MAX ? 'inline' : 'none';
  }

  // Pre-populate selected set from initial checked checkboxes
  document.querySelectorAll('.company-cb:checked').forEach(cb => {
    selected.add(cb.value);
  });
  updateUI();

  // Use the checkbox change event (not click on the row) to avoid
  // the label double-fire issue where selection toggles on then immediately off.
  document.querySelectorAll('.company-cb').forEach(cb => {
    const id  = cb.value;
    const row = document.getElementById('row_' + id);

    // Clicking anywhere on the row should toggle the checkbox
    row.addEventListener('click', (e) => {
      if (cb.disabled) return;
      // If user clicked directly on the checkbox, the browser handles it — skip
      if (e.target === cb) return;
      // Otherwise manually toggle
      if (cb.checked) {
        cb.checked = false;
      } else {
        if (selected.size >= MAX && !selected.has(id)) return;
        cb.checked = true;
      }
      cb.dispatchEvent(new Event('change'));
    });

    cb.addEventListener('change', () => {
      if (cb.checked) {
        if (selected.size >= MAX) {
          cb.checked = false; // Enforce plan limit
          return;
        }
        selected.add(id);
        row.classList.add('selected');
      } else {
        selected.delete(id);
        row.classList.remove('selected');
      }
      updateUI();
    });
  });

  document.getElementById('btnConfirm').addEventListener('click', async () => {
    const btn = document.getElementById('btnConfirm');
    const errBanner = document.getElementById('errorBanner');
    if (selected.size === 0) return;

    // Only send newly selected companies (exclude already-connected disabled ones)
    const newlySelected = Array.from(selected).filter(id => {
      const cb = document.getElementById('cb_' + id);
      return cb && !cb.disabled;
    });

    if (newlySelected.length === 0) {
      errBanner.textContent = 'Please select at least one new company to connect.';
      errBanner.style.display = 'block';
      return;
    }

    btn.disabled = true;
    btn.classList.add('loading');
    btn.innerHTML = '<span class="spinner"></span> Connecting...';
    errBanner.style.display = 'none';

    try {
      const res = await fetch('/api/xero/select-companies', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ selectedTenantIds: newlySelected })
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        // Standardized backend error envelope uses "message"; keep the
        // "error" fallback for backwards compatibility with old responses.
        throw new Error(err.message || err.error || 'Failed to save companies.');
      }

      // Notify parent window and close (mirrors existing SUCCESS_HTML pattern)
      if (window.opener) {
        window.opener.postMessage('xero_connected', '*');
      }
      // For Office dialog context
      if (typeof Office !== 'undefined' && Office.context && Office.context.ui) {
        Office.context.ui.messageParent('xero_connected');
      }
      window.close();
    } catch (err) {
      errBanner.textContent = err.message;
      errBanner.style.display = 'block';
      btn.disabled = false;
      btn.classList.remove('loading');
      btn.innerHTML = 'Connect Selected Companies';
    }
  });
</script>
</body>
</html>`);
        } catch (err) {
            const details = JSON.stringify(err.response?.data || err.message);
            next(new ValidationError('Failed to connect Xero. Please try again.', details));
        }
    };

    /**
     * POST /api/xero/select-companies
     * Receives the user's chosen tenant IDs (from the selection page),
     * reads the pending tokens from session, and persists only the chosen orgs.
     */
    selectCompanies = async (req, res, next) => {
        try {
            const { selectedTenantIds } = req.body;

            if (!selectedTenantIds || !Array.isArray(selectedTenantIds) || selectedTenantIds.length === 0) {
                throw new ValidationError('No companies selected.');
            }

            // Read pending data from session
            const tokens     = req.session?.xero_pending_tokens;
            const tenants    = req.session?.xero_pending_tenants;
            const mail       = req.session?.xero_pending_mail || req.session?.user_mail || null;
            const sessionInfo = JSON.stringify(req.session || {});

            if (!tokens || !tenants) {
                throw new ValidationError('Session expired. Please reconnect Xero.');
            }

            const { XeroToken } = require('../../core/database');
            const whereClause = mail ? { mail } : {};
            const otherCount = await XeroToken.count({
                where: {
                    ...whereClause,
                    tenant_id: {
                        [require('sequelize').Op.notIn]: selectedTenantIds
                    }
                }
            });

            const maxAllowed = req.session?.xero_max_allowed || 10;
            if (otherCount + selectedTenantIds.length > maxAllowed) {
                throw new ValidationError(
                    `Your plan allows a maximum of ${maxAllowed} Xero companies. You currently have ${otherCount} connected companies and selected ${selectedTenantIds.length} more.`
                );
            }

            await XeroService.saveSelectedTenants(selectedTenantIds, tokens, tenants, mail, sessionInfo);

            // Clear pending session data now that we've saved
            delete req.session.xero_pending_tokens;
            delete req.session.xero_pending_tenants;
            delete req.session.xero_pending_mail;

            return res.json({ success: true, connected: selectedTenantIds.length });
        } catch (err) {
            next(err);
        }
    };

    /**
     * POST /api/xero/disconnect
     * Clears the authenticated user's own stored Xero tokens only.
     */
    disconnectXero = async (req, res, next) => {
        try {
            await XeroTokenRepository.clearTokens(req.user.email);
            res.json({ success: true, message: 'Xero tokens cleared successfully.' });
        } catch (error) {
            next(error);
        }
    };

    /**
     * GET /api/xero/tokens
     * Returns the authenticated user's own stored Xero OAuth tokens.
     */
    listXeroTokens = async (req, res, next) => {
        try {
            const tokens = await XeroTokenRepository.getAllTokens(req.user.email);
            res.json({ success: true, tokens });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/xero/contacts
     * Returns a list of mapped ContactDTOs for the authenticated user's tenants.
     */
    getContacts = async (req, res, next) => {
        try {
            const contacts = await XeroService.getContacts(req.user.email);
            res.json({ contacts });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/xero/accounts
     * Returns a list of mapped AccountDTOs for the authenticated user's tenants.
     */
    getAccounts = async (req, res, next) => {
        try {
            const accounts = await XeroService.getAccounts(req.user.email);
            res.json({ accounts });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/xero/classes
     * Returns a list of mapped ClassDTOs for the authenticated user's tenants.
     */
    getClasses = async (req, res, next) => {
        try {
            const classes = await XeroService.getClasses(req.user.email);
            res.json({ classes });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/xero/locations
     * Returns a list of mapped LocationDTOs for the authenticated user's tenants.
     */
    getLocations = async (req, res, next) => {
        try {
            const locations = await XeroService.getLocations(req.user.email);
            res.json({ locations });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/xero/organisation
     * Returns organisation info for the authenticated user's tenants.
     */
    getOrganisation = async (req, res, next) => {
        try {
            const organisation = await XeroService.getOrganisation(req.user.email);
            res.json({ organisation });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/xero/connections
     */
    listConnections = async (req, res, next) => {
        try {
            const mail = req.user.email;
            const list = await XeroService.listConnections(mail);
            return res.json(list);
        } catch (err) {
            return next(err);
        }
    };

    /**
     * GET /api/xero/connections/stats
     */
    getConnectionStats = async (req, res, next) => {
        try {
            const mail = req.user.email;
            const plan = req.query.plan || 'pro';

            const stats = {
                plan: plan.toLowerCase(),
                maxPerPlatform: 10,
                xero:       { connected: 0, remaining: 10 }
            };

            if (plan === 'trial')    stats.maxPerPlatform = 1;
            else if (plan === 'basic')    stats.maxPerPlatform = 1;
            else if (plan === 'standard') stats.maxPerPlatform = 3;

            const xeroStats = await XeroService.getConnectionStats(mail, plan);
            stats.xero = {
                connected: xeroStats.connected,
                remaining: xeroStats.remaining
            };

            return res.json(stats);
        } catch (err) {
            return next(err);
        }
    };

    /**
     * DELETE /api/xero/connections/:id
     */
    disconnectConnection = async (req, res, next) => {
        try {
            const companyId = req.params.id;
            const success = await XeroService.disconnectConnection(companyId, req.user.email);
            return res.json({ success: !!success });
        } catch (err) {
            return next(err);
        }
    };

    /**
     * POST /api/xero/connections/:id/activate
     */
    activateConnection = async (req, res, next) => {
        try {
            const companyId = req.params.id;
            const success = await XeroService.activateConnection(companyId, req.user.email);
            return res.json({ success: !!success });
        } catch (err) {
            return next(err);
        }
    };

    /**
     * PATCH /api/xero/connections/:id/rename
     */
    renameConnection = async (req, res, next) => {
        try {
            const companyId = req.params.id;
            const { companyName } = req.body;
            if (!companyName) {
                throw new ValidationError('companyName is required.');
            }

            const success = await XeroService.renameConnection(companyId, req.user.email, companyName);
            return res.json({ success: !!success });
        } catch (err) {
            return next(err);
        }
    };

    /**
     * GET /api/xero/pull-master-data?companyId=...&tier=...
     */
    pullMasterData = async (req, res, next) => {
        try {
            const { companyId, tier } = req.query;

            const aggregated = await XeroService.pullMasterData(companyId, tier, req.user.email);

            if (!aggregated) {
                const { AppError } = require('../../core/errors/AppError');
                throw new AppError('The requested resource was not found.', 404, 'ERR_NOT_FOUND', `No active connections found for xero.`);
            }

            return res.json({
                company:   aggregated.company.length === 1 ? aggregated.company[0] : aggregated.company,
                customers: aggregated.customers,
                vendors:   aggregated.vendors,
                accounts:  aggregated.accounts,
                classes:   aggregated.classes,
                locations: aggregated.locations,
                isFirstSync: aggregated.isFirstSync
            });
        } catch (err) {
            return next(err);
        }
    };
}

module.exports = new XeroController();