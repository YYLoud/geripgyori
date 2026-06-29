// OneDrive upload helper — MSAL.js + Microsoft Graph API
// App: Geri Teaching Tools | Tenant: SHINKONG WU HO-SU MEMORIAL HOSPITAL

const _OD_CONFIG = {
  clientId: 'c949038d-a9e5-4f00-995b-486faa9ef71b',
  tenantId: '685b9a1c-227e-4acf-be0a-428e5ba2438e',
  folder:   '老年醫學科',           // OneDrive 目標資料夾
  scopes:   ['Files.ReadWrite', 'User.Read']
};

const _msalConfig = {
  auth: {
    clientId:    _OD_CONFIG.clientId,
    authority:   `https://login.microsoftonline.com/${_OD_CONFIG.tenantId}`,
    redirectUri: window.location.origin + window.location.pathname
  },
  cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
};

let _msalInstance = null;
let _account       = null;

async function _initMsal() {
  if (_msalInstance) return _msalInstance;
  _msalInstance = new msal.PublicClientApplication(_msalConfig);
  await _msalInstance.initialize();
  const accounts = _msalInstance.getAllAccounts();
  if (accounts.length > 0) _account = accounts[0];
  return _msalInstance;
}

async function _getToken() {
  const app = await _initMsal();
  const req = { scopes: _OD_CONFIG.scopes, account: _account };
  try {
    const r = await app.acquireTokenSilent(req);
    _account = r.account;
    return r.accessToken;
  } catch {
    const r = await app.acquireTokenPopup(req);
    _account = r.account;
    return r.accessToken;
  }
}

async function _uploadBlob(blob, filename) {
  const token = await _getToken();
  const path  = encodeURIComponent(`${_OD_CONFIG.folder}/${filename}`);
  const res   = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${path}:/content`,
    {
      method:  'PUT',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type':  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      },
      body: blob
    }
  );
  if (!res.ok) {
    const e = await res.json().catch(() => ({}));
    throw new Error(e.error?.message || `HTTP ${res.status}`);
  }
  return res.json();
}

// ── Public API ───────────────────────────────────────────────
window.ODHelper = {
  _blob: null,
  _filename: null,

  // Call after blob is generated; shows & enables the upload button
  setBlob(blob, filename) {
    this._blob     = blob;
    this._filename = filename;
    const btn = document.getElementById('od-upload-btn');
    if (!btn) return;
    btn.style.display = 'flex';
    btn.disabled = false;
    btn.dataset.state = 'ready';
    btn.querySelector('.od-label').textContent = '上傳到 OneDrive';
  },

  async upload() {
    if (!this._blob) return;
    const btn = document.getElementById('od-upload-btn');
    if (!btn) return;

    btn.disabled = true;
    btn.querySelector('.od-label').textContent = '連線中…';

    try {
      await _uploadBlob(this._blob, this._filename);
      btn.querySelector('.od-label').textContent =
        `✓ 已上傳至 OneDrive/${_OD_CONFIG.folder}/`;
      btn.dataset.state = 'done';
      btn.style.background = '#0F6E56';
      btn.style.borderColor = '#0F6E56';
    } catch (err) {
      btn.querySelector('.od-label').textContent = '上傳失敗：' + err.message;
      btn.dataset.state = 'error';
      btn.style.background = '#B91C1C';
      btn.style.borderColor = '#B91C1C';
      btn.disabled = false;
    }
  }
};
