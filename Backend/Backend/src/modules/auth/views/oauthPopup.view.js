'use strict';

const MESSAGE_TYPES = {
    google: { authed: 'google_authed', profile: 'google_profile', cancelled: 'google_cancelled', error: 'google_error', label: 'Google' },
    microsoft: { authed: 'microsoft_authed', profile: 'ms_profile', cancelled: 'ms_cancelled', error: 'microsoft_error', label: 'Microsoft' }
};

function messageTypesFor(provider) {
    const types = MESSAGE_TYPES[provider];
    if (!types) throw new Error(`OAuthPopupView: unknown provider "${provider}"`);
    return types;
}

class OAuthPopupView {
    static renderAuthSuccess({ provider, email, name, token, refreshToken }) {
        const { profile } = messageTypesFor(provider);
        return `<!DOCTYPE html>
<html>
  <head>
    <title>Authentication Successful</title>
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  </head>
  <body style="background:#f8fafc; font-family:sans-serif; text-align:center; padding:50px; color:#172b56;">
    <svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="#2459dd" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="margin-bottom:20px;">
      <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
      <polyline points="22 4 12 14.01 9 11.01"></polyline>
    </svg>
    <h2 style="margin-top:0;">Authentication Successful</h2>
    <p style="color:#64748b;">Returning you to FinAccrual...</p>
    <script>
      var payload = { type: '${profile}', email: '${email}', name: '${name}', token: '${token}', refreshToken: '${refreshToken || ''}' };
      setTimeout(function() {
        if (window.opener) { window.opener.postMessage(payload, '*'); window.close(); }
        else if (typeof Office !== 'undefined' && Office.context && Office.context.ui) { Office.context.ui.messageParent(JSON.stringify(payload)); }
        else { window.close(); }
      }, 800);
    </script>
  </body>
</html>`;
    }

    static renderError({ provider, message }) {
        const { error } = messageTypesFor(provider);
        const safeMessage = String(message || 'Unknown error').replace(/'/g, "\\'");
        return `<!DOCTYPE html>
<html>
  <body style="background:#1a0000;color:#ffaaaa;font-family:sans-serif;text-align:center;padding:40px;">
    <div style="font-size:24px;">❌ Sign-in failed</div>
    <div style="margin-top:10px;font-size:13px;">${safeMessage}</div>
    <script>
      setTimeout(function() {
        if (window.opener) { window.opener.postMessage({ type: '${error}', message: '${safeMessage}' }, '*'); window.close(); }
        else if (typeof Office !== 'undefined') { Office.onReady(function() { Office.context.ui.messageParent(JSON.stringify({ type: '${error}' })); }); }
        else { window.close(); }
      }, 2000);
    </script>
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  </body>
</html>`;
    }
}
module.exports = OAuthPopupView;
