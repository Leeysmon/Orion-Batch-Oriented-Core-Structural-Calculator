(function () {
  const cfg = window.ORION_CONFIG || {};
  const DEV = 'Dr. L. Hulijeli';
  const KEY = 'orion.lastSeen';
  let expired = false;

  if (cfg.expires) {
    // Use the latest date ever seen so winding the system clock back does not unlock the app.
    let now = Date.now();
    try {
      const seen = Number(localStorage.getItem(KEY)) || 0;
      if (seen > now) now = seen; else localStorage.setItem(KEY, String(now));
    } catch (e) { /* storage unavailable */ }
    const [y, m, d] = cfg.expires.split('-').map(Number);
    expired = now >= new Date(y, m - 1, d + 1).getTime();
  }

  window.OrionLock = { expired };

  document.addEventListener('DOMContentLoaded', () => {
    const v = document.getElementById('verText');
    if (v) v.textContent = `Version ${cfg.version} · ${cfg.year}`;
    const ex = document.getElementById('lockText');
    if (ex) ex.textContent = cfg.expires ? `Licence valid until 31 December ${cfg.year}.` : 'No expiry.';
    if (!expired) return;
    document.body.innerHTML = `<div style="min-height:100vh;display:flex;align-items:center;justify-content:center;background:#08121f;color:#e3eefc;font-family:Inter,'Segoe UI',sans-serif;padding:24px">
      <div style="max-width:520px;text-align:center;background:#0f1f35;border:1px solid #25405f;border-radius:18px;padding:32px">
        <h1 style="margin:0 0 10px;font-size:1.4rem;color:#4da3ff">Orion – Access expired</h1>
        <p style="line-height:1.6;color:#8fa9c8">This version (${cfg.version}) was licensed for use until 31 December ${cfg.year} and is now locked.</p>
        <p style="line-height:1.6">Please contact the developer, <strong>${DEV}</strong>, to obtain permission and an updated version.</p>
      </div></div>`;
  });
})();
