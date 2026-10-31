(() => {
  const $ = (id) => document.getElementById(id);
  const logEl = $('log');
  let selected = null;
  let appsMode = null; // 'expired' | 'all'

  function log(msg) {
    const line = `[${new Date().toLocaleTimeString()}] ${msg}`;
    logEl.textContent = (logEl.textContent ? logEl.textContent + '\n' : '') + line;
    logEl.scrollTop = logEl.scrollHeight;
  }

  function setMsg(text, ok) {
    const el = $('actionMsg');
    el.textContent = text || '';
    el.className = ok === true ? 'ok' : ok === false ? 'err' : 'muted';
  }

  async function api(path, opts = {}) {
    const res = await fetch(path, {
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json', ...(opts.headers || {}) },
      ...opts,
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || res.statusText || 'Request failed');
    return data;
  }

  function setConnected(on, status) {
    $('btnConnect').disabled = !!on;
    $('btnDisconnect').disabled = !on;
    $('btnFindExpired').disabled = !on;
    $('btnLoadApps').disabled = !on;
    $('appSearch').disabled = !on;
    $('status').textContent = status || (on ? 'Status: Connected' : 'Status: Disconnected');
    $('footerTenant').textContent = on
      ? `Tenant: ${status.replace(/^Status:\s*/, '')}`
      : 'Tenant: Not connected';
    if (!on) {
      selected = null;
      renderSelected();
      renderApps([]);
      $('secretText').value = '';
      $('btnCopySecret').disabled = true;
    }
  }

  function renderSelected() {
    if (!selected) {
      $('selectedApp').textContent = 'None';
      $('btnGenerate').disabled = true;
      $('btnDeleteExpired').disabled = true;
      $('btnAddAtr').disabled = true;
      return;
    }
    $('selectedApp').textContent = selected.displayName + (selected.endDate ? ` · expired ${new Date(selected.endDate).toLocaleDateString()}` : '');
    $('btnGenerate').disabled = false;
    $('btnDeleteExpired').disabled = appsMode !== 'expired';
    $('btnAddAtr').disabled = false;
  }

  function renderApps(apps) {
    const tbody = $('appsTable').querySelector('tbody');
    tbody.innerHTML = '';
    if (!apps.length) {
      const tr = document.createElement('tr');
      tr.innerHTML = '<td colspan="2" class="muted">No applications</td>';
      tbody.appendChild(tr);
      return;
    }
    for (const a of apps) {
      const tr = document.createElement('tr');
      if (selected && selected.applicationId === a.applicationId) tr.classList.add('selected');
      const detail = a.endDate
        ? `Oldest expired: ${new Date(a.endDate).toLocaleDateString()} (${a.expiredCount || 1})`
        : (a.appId || '');
      tr.innerHTML = `<td>${escapeHtml(a.displayName || '')}</td><td class="muted">${escapeHtml(detail)}</td>`;
      tr.addEventListener('click', () => {
        selected = a;
        renderSelected();
        [...tbody.querySelectorAll('tr')].forEach((r) => r.classList.remove('selected'));
        tr.classList.add('selected');
        $('secretText').value = '';
        $('btnCopySecret').disabled = true;
        setMsg('');
      });
      tbody.appendChild(tr);
    }
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  async function loadTenants() {
    const data = await api('/api/tenants');
    const sel = $('tenantSelect');
    sel.innerHTML = '';
    for (const t of data.tenants || []) {
      const opt = document.createElement('option');
      opt.value = t.tenantId;
      opt.textContent = t.displayText || t.tenantId;
      sel.appendChild(opt);
    }
    log(`Loaded ${(data.tenants || []).length} tenant(s) from shared XOA store.`);
  }

  async function refreshStatus() {
    try {
      const s = await api('/api/status');
      if (s.connected) {
        const label = s.orgName ? `${s.orgName} (${s.tenantId})` : s.tenantId;
        setConnected(true, `Status: Connected · ${label}`);
      } else {
        setConnected(false);
      }
    } catch {
      setConnected(false);
    }
  }

  $('btnRefreshTenants').addEventListener('click', () => {
    loadTenants().catch((e) => { log(e.message); setMsg(e.message, false); });
  });

  $('btnConnect').addEventListener('click', async () => {
    try {
      setMsg('Connecting…');
      const tenantId = $('tenantSelect').value;
      if (!tenantId) throw new Error('Select a tenant');
      const data = await api('/api/connect', { method: 'POST', body: JSON.stringify({ tenantId }) });
      const label = data.orgName ? `${data.orgName} (${data.tenantId})` : data.tenantId;
      setConnected(true, `Status: Connected · ${label}`);
      setMsg('Connected.', true);
      log(`Connected to ${label}`);
    } catch (e) {
      setMsg(e.message, false);
      log(`Connect failed: ${e.message}`);
    }
  });

  $('btnDisconnect').addEventListener('click', async () => {
    try {
      await api('/api/disconnect', { method: 'POST', body: '{}' });
      setConnected(false);
      setMsg('Disconnected.');
      log('Disconnected.');
    } catch (e) {
      setMsg(e.message, false);
    }
  });

  $('btnFindExpired').addEventListener('click', async () => {
    try {
      setMsg('Finding expired secrets…');
      $('appsHint').textContent = 'Loading expired secrets…';
      const data = await api('/api/apps/expired');
      appsMode = 'expired';
      selected = null;
      renderSelected();
      renderApps(data.apps || []);
      $('appsHint').textContent = `Found ${(data.apps || []).length} application(s) with expired secrets. Click a row to select.`;
      setMsg(`Found ${(data.apps || []).length} app(s).`, true);
      log(`Expired secrets: ${(data.apps || []).length}`);
    } catch (e) {
      setMsg(e.message, false);
      log(e.message);
    }
  });

  async function loadAllApps() {
    const q = $('appSearch').value.trim();
    const qs = q ? `?q=${encodeURIComponent(q)}` : '';
    setMsg('Loading applications…');
    const data = await api('/api/apps' + qs);
    appsMode = 'all';
    selected = null;
    renderSelected();
    renderApps(data.apps || []);
    $('appsHint').textContent = `${(data.apps || []).length} application(s). Click a row to select.`;
    setMsg(`Loaded ${(data.apps || []).length} app(s).`, true);
  }

  $('btnLoadApps').addEventListener('click', () => {
    loadAllApps().catch((e) => { setMsg(e.message, false); log(e.message); });
  });

  let searchTimer = null;
  $('appSearch').addEventListener('input', () => {
    if (appsMode !== 'all') return;
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      loadAllApps().catch((e) => setMsg(e.message, false));
    }, 300);
  });

  $('btnGenerate').addEventListener('click', async () => {
    if (!selected) return;
    try {
      setMsg('Generating secret…');
      const data = await api('/api/secrets/generate', {
        method: 'POST',
        body: JSON.stringify({ applicationId: selected.applicationId }),
      });
      $('secretText').value = data.secretText || '';
      $('btnCopySecret').disabled = !data.secretText;
      setMsg(`Secret created (${data.displayName}). Copy it now; it is shown only once.`, true);
      log(`Generated secret for ${selected.displayName}: ${data.displayName}`);
    } catch (e) {
      setMsg(e.message, false);
      log(e.message);
    }
  });

  $('btnDeleteExpired').addEventListener('click', async () => {
    if (!selected) return;
    if (!confirm(`Delete oldest expired secret for "${selected.displayName}"?`)) return;
    try {
      setMsg('Deleting expired secret…');
      await api('/api/secrets/delete-expired', {
        method: 'POST',
        body: JSON.stringify({ applicationId: selected.applicationId }),
      });
      setMsg('Expired secret deleted.', true);
      log(`Deleted expired secret for ${selected.displayName}`);
      if (appsMode === 'expired') $('btnFindExpired').click();
    } catch (e) {
      setMsg(e.message, false);
      log(e.message);
    }
  });

  $('btnAddAtr').addEventListener('click', async () => {
    if (!selected) return;
    if (!confirm(`Add Barracuda XDR ATR permissions to "${selected.displayName}" and grant admin consent?`)) return;
    try {
      setMsg('Adding ATR permissions…');
      const data = await api('/api/atr/add', {
        method: 'POST',
        body: JSON.stringify({ applicationId: selected.applicationId }),
      });
      const parts = [];
      if ((data.permissionsAdded || []).length) parts.push('added: ' + data.permissionsAdded.join(', '));
      if ((data.consentGranted || []).length) parts.push('consent: ' + data.consentGranted.join(', '));
      if ((data.consentFailed || []).length) parts.push('consent failed: ' + data.consentFailed.join('; '));
      const ok = !(data.consentFailed || []).length;
      setMsg(parts.join(' · ') || 'Done.', ok);
      log(`ATR for ${selected.displayName}: ${parts.join(' · ') || 'ok'}`);
    } catch (e) {
      setMsg(e.message, false);
      log(e.message);
    }
  });

  $('btnCopySecret').addEventListener('click', async () => {
    const t = $('secretText').value;
    if (!t) return;
    try {
      await navigator.clipboard.writeText(t);
      setMsg('Secret copied to clipboard. Clear clipboard after pasting.', true);
    } catch {
      $('secretText').select();
      setMsg('Select and copy manually (clipboard API blocked).', false);
    }
  });

  loadTenants()
    .then(refreshStatus)
    .catch((e) => { log(e.message); setMsg(e.message, false); });
})();
