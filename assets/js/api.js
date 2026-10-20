/* Komunikasi ke backend Apps Script. POST + text/plain agar tidak memicu preflight CORS. */
const API = {
  token() { try { return sessionStorage.getItem('ev_token') || ''; } catch (e) { return ''; } },
  setToken(t) { try { t ? sessionStorage.setItem('ev_token', t) : sessionStorage.removeItem('ev_token'); } catch (e) {} },
  async call(action, data, opts) {
    opts = opts || {};
    let res;
    try {
      res = await fetch(CONFIG.API_URL, {
        method: 'POST', redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action, data: data || {}, token: opts.noAuth ? '' : API.token() })
      });
    } catch (e) { throw { code: 'NETWORK', message: 'Tidak dapat terhubung ke server. Periksa koneksi internet dan API_URL di config.js.' }; }
    let j;
    try { j = await res.json(); } catch (e) { throw { code: 'BAD_RESPONSE', message: 'Respons server tidak valid. Pastikan deployment diatur "Anyone" dan memakai URL /exec.' }; }
    if (!j.ok) { if (j.error && j.error.code === 'UNAUTHORIZED') API.setToken(''); throw j.error || { code: 'UNKNOWN', message: 'Terjadi kesalahan.' }; }
    return j.data;
  }
};
