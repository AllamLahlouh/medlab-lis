/* =============================================
   MedLab LIS v5 — Application Logic
   All data is stored in localStorage (no server needed)
   ============================================= */

// ═══════════════════════════════════════
// CONSTANTS & STATE
// ═══════════════════════════════════════
const ADMIN = { u: 'Admin', p: 'Admin12345' };
const SK = { users: 'lisV5_users', patients: 'lisV5_patients', tests: 'lisV5_tests', orders: 'lisV5_orders', results: 'lisV5_results', billing: 'lisV5_billing', shifts: 'lisV5_shifts', curShift: 'lisV5_curShift' };
let CU = null, loginRole = 'doctor';
let pendingAuthCb = null;


// ═══════════════════════════════════════
// SUPABASE STORAGE
// ═══════════════════════════════════════
const _sb = supabase.createClient('https://tuwwdoyxrwoeqafmpkeg.supabase.co', 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InR1d3dkb3l4cndvZXFhZm1wa2VnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzM5ODY1NTUsImV4cCI6MjA4OTU2MjU1NX0.km8-p7wCgbw_pl6ADQZfQKJRpwF_x0AwYOf_0_hX1Hg');
let DB = { users: [], patients: [], tests: [], orders: [], results: [], billing: [], shifts: [] };
let curShift = null;

function ld(k) { return DB[k] || []; }
function ldOne(k) { if (k === 'curShift') return curShift; return DB[k]; }

// Sync data to Supabase (Atomic: only upsert/save the modified item)
async function sv(table, item) {
  if (!item || !item.id) return;
  // Update local cache
  const idx = DB[table].findIndex(x => x.id === item.id);
  if (idx >= 0) DB[table][idx] = item; else DB[table].push(item);
  syncMsg();
  try {
    const { error } = await _sb.from(table).upsert(item);
    if (error) throw error;
  } catch (e) { console.error(`Supabase Sync Error [${table}]:`, e); toast('Sync Error: failed to save to cloud', 'error'); }
}

async function svOne(k, d) {
  if (k === 'lisV5_curShift') curShift = d;
  DB[k] = d;
  syncMsg();
  try {
    const { error } = await _sb.from('config').upsert({ key: k, value: d, updated_at: new Date().toISOString() });
    if (error) throw error;
  } catch (e) { console.error('Supabase Sync Error [config]:', e); toast('Sync Error: failed to save config to cloud', 'error'); }
}

async function delRow(table, id) {
  DB[table] = DB[table].filter(x => x.id !== id);
  try {
    const { error } = await _sb.from(table).delete().eq('id', id);
    if (error) throw error;
  } catch (e) { console.error(`Supabase Sync Error [delete ${table}]:`, e); toast('Sync Error: failed to delete from cloud', 'error'); }
}

function syncMsg() {
  const e = document.getElementById('syncInfo');
  if (e) { e.innerHTML = '● Supabase Synced ' + new Date().toLocaleTimeString('en-GB'); e.style.color = 'var(--green2)'; }
}

// ═══════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════
function tod() { return new Date().toISOString().slice(0, 10); }
function nowT() { return new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }); }
function fmtD(d) { if (!d || d === '—') return '—'; try { return new Date(d + 'T00:00:00').toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }); } catch { return d; } }
function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function gV(id) { const e = document.getElementById(id); return e ? e.value : ''; }
function sV(id, v) { const e = document.getElementById(id); if (e) e.value = (v == null ? '' : v); }
function genID(pre, arr) { const ns = arr.map(x => parseInt((x.id || '').replace(/\D/g, '')) || 0); return pre + String(Math.max(0, ...ns) + 1).padStart(4, '0'); }
function toast(msg, t = 'success') {
  const el = document.getElementById('toast');
  el.textContent = msg; el.style.background = t === 'error' ? 'var(--red2)' : t === 'warn' ? '#795548' : 'var(--green2)';
  el.classList.add('show'); setTimeout(() => el.classList.remove('show'), 3200);
}
function openM(id) { document.getElementById(id).classList.add('open'); }
function closeM(id) { document.getElementById(id).classList.remove('open'); }
document.querySelectorAll('.mOver').forEach(m => { m.addEventListener('click', function (e) { if (e.target === this) this.classList.remove('open'); }); });
function conf2(title, msg, cb) {
  document.getElementById('confTitle').textContent = title;
  document.getElementById('confMsg').textContent = msg;
  const b = document.getElementById('confOK'); b.onclick = () => { closeM('confModal'); cb(); };
  openM('confModal');
}
function sBadge(s) { const m = { Completed: 'bG', 'In Progress': 'bGd', Pending: 'bGr', Cancelled: 'bR' }; return `<span class="badge ${m[s] || 'bGr'}">${esc(s)}</span>`; }
function pBadge(p) { const m = { Routine: 'bB', Urgent: 'bO', STAT: 'bR' }; return `<span class="badge ${m[p] || 'bB'}">${p === 'STAT' ? '⚡ ' : p === 'Urgent' ? '🔴 ' : ''}${esc(p)}</span>`; }
function rBadge(r) { if (r === 'doctor') return `<span class="rBadge rbD">👨‍⚕️ Doctor</span>`; if (r === 'lab') return `<span class="rBadge rbL">🔬 Lab Tech</span>`; return `<span class="rBadge rbA">🛡️ Admin</span>`; }
function getTests(o) { if (Array.isArray(o.tests)) return o.tests; const t = [];['test1', 'test2', 'test3', 'test4', 'test5', 'test6'].forEach(k => { if (o[k]) t.push(o[k]); }); return t; }
function show(id, v) { const e = document.getElementById(id); if (!e) return; v ? e.classList.remove('hn') : e.classList.add('hn'); }

// ═══════════════════════════════════════
// INIT DATA
// ═══════════════════════════════════════
async function initData() {
  const tables = ['users', 'patients', 'tests', 'orders', 'results', 'billing', 'shifts', 'config'];
  try {
    for (const t of tables) {
      const { data, error } = await _sb.from(t).select('*');
      if (error) throw error;

      if (data && data.length > 0) {
        if (t === 'config') {
          data.forEach(row => {
            DB[row.key] = typeof row.value === 'string' ? JSON.parse(row.value) : row.value;
          });
          curShift = DB['lisV5_curShift'] || null;
        } else {
          DB[t] = data;
        }
      } else if (t !== 'config') {
        // Migration from localStorage if cloud is empty
        const migratedKey = 'lisV5_migrated_' + t;
        if (!localStorage.getItem(migratedKey)) {
          const raw = localStorage.getItem(SK[t]);
          if (raw) {
            const arr = JSON.parse(raw);
            if (arr && arr.length > 0) {
              DB[t] = arr;
              await _sb.from(t).upsert(arr);
            }
          }
          localStorage.setItem(migratedKey, '1');
        }
      }
    }

    const { data: cdata } = await _sb.from('config').select('*');
    if (cdata && cdata.length > 0) {
      for (const row of cdata) {
        DB[row.key] = typeof row.value === 'string' ? JSON.parse(row.value) : row.value;
      }
      curShift = DB['lisV5_curShift'] || null;
    } else {
      // BUG FIX: Only migrate config once
      const configMigrated = localStorage.getItem('lisV5_configMigrated');
      if (!configMigrated) {
        for (let i = 0; i < localStorage.length; i++) {
          const k = localStorage.key(i);
          if (k && (k.startsWith('lisV5_hist') || k === 'lisV5_curShift')) {
            try {
              const val = JSON.parse(localStorage.getItem(k));
              DB[k] = val;
              if (k === 'lisV5_curShift') curShift = val;
              await _sb.from('config').upsert({ key: k, value: JSON.stringify(val) });
            } catch { }
          }
        }
        localStorage.setItem('lisV5_configMigrated', '1');
      }
    }

    // Global Realtime Subscription — BUG FIX: also re-render labDash and users views
    _sb.channel('public-db').on('postgres_changes', { event: '*', schema: 'public' }, async (payload) => {
      const t = payload.table;
      if (t === 'config') {
        const { data } = await _sb.from('config').select('*');
        if (data) {
          for (const row of data) { DB[row.key] = typeof row.value === 'string' ? JSON.parse(row.value) : row.value; }
          curShift = DB['lisV5_curShift'] || null;
        }
      } else {
        const { data } = await _sb.from(t).select('*');
        if (data) DB[t] = data;
      }
      // Re-render whichever view is currently active
      if (typeof renderPatients === 'function' && document.getElementById('patients-view')?.classList.contains('active')) renderPatients();
      if (typeof renderOrders === 'function' && document.getElementById('orders-view')?.classList.contains('active')) renderOrders();
      if (typeof renderResults === 'function' && document.getElementById('results-view')?.classList.contains('active')) renderResults();
      if (typeof renderTests === 'function' && document.getElementById('tests-view')?.classList.contains('active')) renderTests();
      if (typeof renderBilling === 'function' && document.getElementById('billing-view')?.classList.contains('active')) renderBilling();
      if (typeof renderAdmDash === 'function' && document.getElementById('admDash-view')?.classList.contains('active')) renderAdmDash();
      if (typeof renderLabDash === 'function' && document.getElementById('labDash-view')?.classList.contains('active')) renderLabDash();
      if (typeof renderUsers === 'function' && document.getElementById('users-view')?.classList.contains('active')) renderUsers();
      if (typeof renderDocDash === 'function' && document.getElementById('docDash-view')?.classList.contains('active')) renderDocDash();
    }).subscribe();
  } catch (e) {
    console.error('initData error:', e);
    toast('Failed to load data from Supabase. Check your connection.', 'error');
  }
}

async function forceRefresh() {
  const btn = document.getElementById('refreshBtn');
  const btn2 = document.getElementById('refreshBtnSide');
  if (btn) btn.classList.add('loading');
  if (btn2) btn2.classList.add('loading');
  toast('Refreshing data...', 'warn');
  await initData();
  const page = document.querySelector('.view.active')?.id.replace('-view', '');
  if (page && typeof nav === 'function') nav(page);
  syncMsg();
  if (btn) btn.classList.remove('loading');
  if (btn2) btn2.classList.remove('loading');
  toast('Data updated');
}

// ═══════════════════════════════════════
// AUTH
// ═══════════════════════════════════════
function setRole(r) {
  loginRole = r;
  document.getElementById('ltA').classList.toggle('active', r === 'admin');
  document.getElementById('ltD').classList.toggle('active', r === 'doctor');
  document.getElementById('ltL').classList.toggle('active', r === 'lab');
  sV('lUser', ''); sV('lPass', '');
  document.getElementById('lErr').classList.remove('show');
  const btn = document.getElementById('lBtn');
  btn.className = 'lBtn ' + (r === 'admin' ? 'bA' : r === 'doctor' ? 'bD' : 'bL');
  const hintTitles = { admin: '🔑 Admin Credentials', doctor: '🔑 Doctor Login', lab: '🔑 Lab Technician Login' };
  document.getElementById('lHintTitle').textContent = hintTitles[r];
  const hints = { admin: 'Admin: Username <b>Admin</b> / Password <b>**********</b>', doctor: 'Doctor accounts are created by the Admin panel. Use credentials provided by your admin.', lab: 'Lab technician accounts are created by the Admin panel. Use credentials provided by your admin.' };
  document.getElementById('lHintBody').innerHTML = hints[r];
}

function doLogin() {
  const u = gV('lUser').trim().toLowerCase();
  const p = gV('lPass');
  const err = document.getElementById('lErr'); err.classList.remove('show');

  // Auto-detect Admin login
  if (u === ADMIN.u.toLowerCase() && p === ADMIN.p) {
    CU = { id: 'admin', username: 'Admin', role: 'admin', name: 'System Administrator', avatar: '🛡️' };
    finishLogin();
    return;
  }

  if (loginRole !== 'admin') {
    const found = ld('users').find(x => x.username.toLowerCase() === u && x.password === p && x.role === loginRole && x.status === 'Active');
    if (found) { CU = { ...found, avatar: found.role === 'doctor' ? '👨‍⚕️' : '🔬' }; finishLogin(); return; }
  }

  err.classList.add('show'); sV('lPass', '');
}

function finishLogin() {
  localStorage.setItem('lisV5_user', JSON.stringify(CU));
  document.getElementById('loginScreen').classList.add('hidden');
  document.getElementById('app').classList.remove('hidden');
  document.getElementById('uName').textContent = CU.username;
  document.getElementById('uRole').textContent = CU.role.toUpperCase();
  document.getElementById('uAv').textContent = CU.avatar;
  setDates(); setupUI();
}

function doLogout() {
  localStorage.removeItem('lisV5_user'); CU = null;
  document.getElementById('app').classList.add('hidden');
  document.getElementById('loginScreen').classList.remove('hidden');
  sV('lUser', ''); sV('lPass', ''); setRole('doctor');
}

function checkSession() { const s = localStorage.getItem('lisV5_user'); if (s) { try { CU = JSON.parse(s); finishLogin(); } catch { } } }

function setupUI() {
  const r = CU.role;
  const vb = document.getElementById('vBadge'); vb.textContent = `v5.0 · ${r.toUpperCase()}`; vb.className = `vbadge v${r[0].toUpperCase()}`;
  const rp = document.getElementById('rolePill');
  if (r === 'admin') { rp.textContent = '🛡️ Admin'; rp.style.background = 'var(--purple-lt)'; rp.style.color = 'var(--adm)'; rp.style.border = '1px solid var(--adm)'; }
  else if (r === 'doctor') { rp.textContent = '👨‍⚕️ Doctor'; rp.style.background = 'var(--sky-lt)'; rp.style.color = 'var(--doc)'; rp.style.border = '1px solid var(--doc)'; }
  else { rp.textContent = '🔬 Lab Tech'; rp.style.background = 'var(--teal-lt)'; rp.style.color = 'var(--lab)'; rp.style.border = '1px solid var(--lab)'; }
  show('nAdmS', r === 'admin'); show('nAdmD', r === 'admin'); show('nUsers', r === 'admin');
  show('nDocS', r === 'doctor'); show('nDocD', r === 'doctor');
  show('nLabS', r === 'lab'); show('nLabD', r === 'lab');
  show('nTests', r === 'admin' || r === 'lab');
  show('nRes', r === 'admin' || r === 'lab');
  show('nBill', r === 'admin' || r === 'doctor');
  show('btnExp', true); show('btnImp', true);
  // Only lab can enter results
  const bar = document.getElementById('btnAddResult'); if (bar) { r === 'lab' ? bar.classList.remove('hn') : bar.classList.add('hn'); }
  if (r === 'doctor') { document.getElementById('patTitle').textContent = 'My Patients'; document.getElementById('ordTitle').textContent = 'My Orders'; document.getElementById('nPatsLbl').textContent = 'My Patients'; document.getElementById('nOrdsLbl').textContent = 'My Orders'; }
  else if (r === 'lab') { document.getElementById('nOrdsLbl').textContent = 'Orders Queue'; }
  if (r === 'admin') nav('admDash');
  else if (r === 'doctor') nav('docDash');
  else nav('labDash');
}

// ═══════════════════════════════════════
// NAVIGATION
// ═══════════════════════════════════════
const PT = { 'admDash': 'Admin Dashboard', 'users': 'User Management', 'docDash': 'My Dashboard', 'labDash': 'Lab Dashboard', 'patients': 'Patients', 'tests': 'Tests Catalog', 'orders': 'Lab Orders', 'results': 'Lab Results', 'billing': 'Billing & Invoices', 'report': 'Print Report' };
function nav(page) {
  if (page === 'billing' && CU.role === 'lab') { toast('Lab workers have no billing access', 'error'); return; }
  if (page === 'results' && CU.role === 'doctor') { toast('Results are managed by lab staff', 'warn'); return; }
  if (page === 'results' && CU.role !== 'admin' && CU.role !== 'lab') { toast('Lab access only', 'error'); return; }
  if (page === 'users' && CU.role !== 'admin') { toast('Admin only', 'error'); return; }
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  document.querySelectorAll('.ni').forEach(n => n.classList.remove('active'));
  const vEl = document.getElementById(page + '-view'); if (vEl) vEl.classList.add('active');
  const ni = document.querySelector(`[onclick="nav('${page}')"]`); if (ni) ni.classList.add('active');
  document.getElementById('pageTitle').textContent = PT[page] || page;
  closeSB();
  const fn = { admDash: renderAdmDash, users: renderUsers, docDash: renderDocDash, labDash: renderLabDash, patients: renderPatients, tests: () => { renderTests(); fillTDeptF(); }, orders: renderOrders, results: renderResults, billing: renderBilling, report: populateRepSel };
  if (fn[page]) fn[page]();
}

// ═══════════════════════════════════════
// SIDEBAR MOBILE
// ═══════════════════════════════════════
function toggleSB() { document.getElementById('sidebar').classList.toggle('open'); document.getElementById('sbOverlay').classList.toggle('on'); }
function closeSB() { document.getElementById('sidebar').classList.remove('open'); document.getElementById('sbOverlay').classList.remove('on'); }

// ═══════════════════════════════════════
// DATES
// ═══════════════════════════════════════
function setDates() {
  const d = new Date();
  document.getElementById('topDate').textContent = d.toLocaleDateString('en-GB', { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric' });
  document.getElementById('footDate').textContent = d.toLocaleDateString('en-GB');
}

// ═══════════════════════════════════════
// KPI HELPER
// ═══════════════════════════════════════
function renderKPIs(cid, kpis) {
  document.getElementById(cid).innerHTML = kpis.map(k => `
    <div class="kCard"><div class="kTop"><div class="kIco" style="background:${k.bg};color:${k.col}">${k.ico}</div></div>
    <div class="kVal" style="color:${k.col}">${k.val}</div><div class="kLbl">${k.lbl}</div></div>`).join('');
}

// ═══════════════════════════════════════
// BAR CHART HELPER
// ═══════════════════════════════════════
function renderBar(cid, data, colors, pre = '') {
  const el = document.getElementById(cid); if (!el) return;
  const vals = Object.values(data), mx = Math.max(...vals, 1);
  el.innerHTML = Object.entries(data).map(([l, v], i) => `
    <div class="bI"><div class="bLbl" title="${esc(l)}">${esc(l)}</div>
    <div class="bTr"><div class="bFil" style="width:${Math.round(v / mx * 100)}%;background:${colors[i % colors.length]}"></div></div>
    <div class="bVl">${pre}${v}</div></div>`).join('') || '<div style="color:var(--gray4);font-size:11px">No data</div>';
}

// ═══════════════════════════════════════
// ADMIN DASHBOARD
// ═══════════════════════════════════════
function renderAdmDash() {
  const users = ld('users'), pats = ld('patients'), ords = ld('orders'), bill = ld('billing');
  const rev = bill.reduce((s, x) => s + Number(x.total || 0), 0);
  const pend = ords.filter(o => o.status === 'Pending' || o.status === 'In Progress').length;
  renderKPIs('admKPI', [
    { lbl: 'Total Users', val: users.length, ico: '👥', col: 'var(--adm)', bg: 'var(--purple-lt)' },
    { lbl: 'Total Patients', val: pats.length, ico: '👤', col: 'var(--doc)', bg: 'var(--sky-lt)' },
    { lbl: 'Active Orders', val: pend, ico: '📋', col: 'var(--orange2)', bg: 'var(--orange-lt)' },
    { lbl: 'Revenue (₪)', val: '₪' + rev.toLocaleString(), ico: '💰', col: 'var(--green2)', bg: 'var(--green-lt)' },
  ]);
  const depts = {}; ld('tests').forEach(t => { depts[t.dept] = (depts[t.dept] || 0) + 1; });
  renderBar('deptChart', depts, ['#1565C0', '#00695C', '#4A148C', '#E65100', '#B71C1C', '#1B5E20', '#D32F2F', '#6A1B9A']);
  const revcats = {}; bill.forEach(b => { const n = b.patient.split(' ')[0]; revcats[n] = (revcats[n] || 0) + Number(b.total || 0); });
  renderBar('revChart', revcats, ['#2E7D32', '#1565C0', '#E65100', '#4A148C', '#C62828'], '₪');
  document.getElementById('recentUsers').innerHTML = users.slice(-5).reverse().map(u => `
    <tr><td class="tM">${esc(u.username)}</td><td class="tN">${esc(u.name)}</td><td>${rBadge(u.role)}</td>
    <td><span class="badge ${u.status === 'Active' ? 'bG' : 'bR'}">${esc(u.status)}</span></td>
    <td class="tM">${fmtD(u.created)}</td></tr>`).join('') || '<tr><td colspan="5" class="empty"><p>No users</p></td></tr>';
  document.getElementById('recentOrdsAdm').innerHTML = ords.slice(-5).reverse().map(o => `
    <tr><td class="tM">${esc(o.id)}</td><td class="tN">${esc(o.patient)}</td>
    <td style="font-size:11px">${esc(o.doctor)}</td><td>${sBadge(o.status)}</td></tr>`).join('') || '<tr><td colspan="4" class="empty"><p>No orders</p></td></tr>';
}

// ═══════════════════════════════════════
// DOCTOR DASHBOARD
// ═══════════════════════════════════════
function renderDocDash() {
  const mp = ld('patients').filter(p => p.ownerId === CU.id);
  const mo = ld('orders').filter(o => o.doctorId === CU.id);
  const mb = ld('billing').filter(b => b.doctorId === CU.id);
  const allRes = ld('results').filter(r => mo.find(o => o.id === r.orderId));
  const activeOrds = mo.filter(o => o.status === 'Pending' || o.status === 'In Progress');
  const completedToday = mo.filter(o => o.status === 'Completed' && o.date === tod()).length;
  const outs = mb.reduce((s, b) => s + (Number(b.total || 0) - Number(b.paid || 0)), 0);
  renderKPIs('docKPI', [
    { lbl: 'My Patients', val: mp.length, ico: '👤', col: 'var(--doc)', bg: 'var(--sky-lt)' },
    { lbl: 'Active Orders', val: activeOrds.length, ico: '⏳', col: 'var(--orange2)', bg: 'var(--orange-lt)' },
    { lbl: 'Completed Today', val: completedToday, ico: '✅', col: 'var(--green2)', bg: 'var(--green-lt)' },
    { lbl: 'Total Results', val: allRes.length, ico: '🔬', col: 'var(--teal)', bg: 'var(--teal-lt)' },
  ]);
  // Recent patients
  document.getElementById('drPatTbl').innerHTML = mp.slice(-6).reverse().map(p => `
    <tr><td class="tM">${esc(p.id)}</td><td class="tNC" onclick="openProf('${esc(p.id)}')">${esc(p.name)}</td>
    <td>${p.gender ? `<span class="badge ${p.gender === 'Male' ? 'bB' : 'bP'}">${esc(p.gender)}</span>` : ''}</td>
    <td class="tM">${esc(p.phone || '—')}</td><td class="tM">${fmtD(p.date)}</td></tr>`).join('')
    || '<tr><td colspan="5" class="empty"><span class="eIco">👤</span><p>No patients yet</p></td></tr>';
  // Active (pending/in-progress) orders
  document.getElementById('drOrdTbl').innerHTML = activeOrds.slice(0, 8).map(o => `
    <tr><td class="tM">${esc(o.id)}</td><td class="tN">${esc(o.patient)}</td>
    <td style="font-size:10px">${getTests(o).slice(0, 2).map(t => esc(t)).join(', ')}${getTests(o).length > 2 ? ' …' : ''}</td>
    <td>${pBadge(o.priority || 'Routine')}</td><td>${sBadge(o.status)}</td></tr>`).join('')
    || '<tr><td colspan="5" class="empty"><span class="eIco">✅</span><p>No active orders</p></td></tr>';
  // Recent results table
  document.getElementById('drAbnTbl').innerHTML = allRes.slice(-8).reverse().map(r => `
    <tr><td class="tN">${esc(r.patient)}</td>
    <td style="font-size:11px">${esc(r.testName)}</td>
    <td style="font-size:10px;color:var(--gray5)">${esc(r.paramName || '—')}</td>
    <td style="font-weight:800;font-family:monospace">${esc(r.value)}</td>
    <td style="font-size:10px">${esc(r.unit || '—')}</td>
    <td style="font-size:10px">${esc(r.norm || '—')}</td>
    <td class="tM">${fmtD(r.date)}</td></tr>`).join('')
    || '<tr><td colspan="7" class="empty"><span class="eIco">🔬</span><p>No results yet</p></td></tr>';
  // Orders by status mini-chart
  const statMap = {}; mo.forEach(o => { statMap[o.status] = (statMap[o.status] || 0) + 1; });
  renderBar('drStatusChart', statMap, ['#2E7D32', '#E65100', '#1565C0', '#C62828']);
}

// ═══════════════════════════════════════
// LAB DASHBOARD
// ═══════════════════════════════════════
function renderLabDash() {
  const sh = ldOne('curShift') || { id: '—', labUser: CU.username, labName: CU.name, startDate: tod(), startTime: nowT(), ordersProcessed: 0, resultsEntered: 0, patientsAdded: 0 };
  document.getElementById('shiftLbl').textContent = sh.id;
  document.getElementById('shiftInfo').textContent = `Started: ${fmtD(sh.startDate)} at ${sh.startTime} · Tech: ${sh.labName || sh.labUser}`;
  document.getElementById('shiftStats').innerHTML = `
    <div class="shSt"><div class="sN">${sh.ordersProcessed || 0}</div><div class="sL">Orders Processed</div></div>
    <div class="shSt"><div class="sN">${sh.resultsEntered || 0}</div><div class="sL">Results Entered</div></div>
    <div class="shSt"><div class="sN">${sh.patientsAdded || 0}</div><div class="sL">Patients Registered</div></div>
    <div class="shSt"><div class="sN" style="color:#90CAF9">${sh.startTime || '—'}</div><div class="sL">Shift Start</div></div>`;
  const allOrds = ld('orders');
  const pending = allOrds.filter(o => o.status === 'Pending');
  const inProg = allOrds.filter(o => o.status === 'In Progress');
  const todayRes = ld('results').filter(r => r.date === tod());
  const completedToday = allOrds.filter(o => o.status === 'Completed' && o.date === tod()).length;
  renderKPIs('labKPI', [
    { lbl: 'Pending Orders', val: pending.length, ico: '⏳', col: 'var(--orange2)', bg: 'var(--orange-lt)' },
    { lbl: 'In Progress', val: inProg.length, ico: '🔬', col: 'var(--lab)', bg: 'var(--teal-lt)' },
    { lbl: 'Results Today', val: todayRes.length, ico: '📊', col: 'var(--blue)', bg: 'var(--sky-lt)' },
    { lbl: 'Completed Today', val: completedToday, ico: '✅', col: 'var(--green2)', bg: 'var(--green-lt)' },
  ]);
  // Priority breakdown chart
  const priMap = {}; allOrds.filter(o => o.status !== 'Completed' && o.status !== 'Cancelled').forEach(o => { const p = o.priority || 'Routine'; priMap[p] = (priMap[p] || 0) + 1; });
  renderBar('labPriChart', priMap, ['var(--blue)', 'var(--orange2)', 'var(--red2)']);
  // Dept breakdown of today's results
  const tests = ld('tests');
  const deptMap = {}; todayRes.forEach(r => { const t = tests.find(x => x.id === r.testId || x.name === r.testName); const d = t?.dept || 'Other'; deptMap[d] = (deptMap[d] || 0) + 1; });
  renderBar('labDeptChart', deptMap, ['#1565C0', '#00695C', '#4A148C', '#E65100', '#B71C1C', '#1B5E20', '#D32F2F']);
  // Pending orders table
  document.getElementById('labPendOrds').innerHTML = allOrds.filter(o => o.status !== 'Completed' && o.status !== 'Cancelled').sort((a, b) => { const pw = { STAT: 0, Urgent: 1, Routine: 2 }; return (pw[a.priority] || 2) - (pw[b.priority] || 2); }).slice(0, 12).map(o => `
    <tr><td class="tM">${esc(o.id)}</td><td class="tN">${esc(o.patient)}</td><td style="font-size:11px">${esc(o.doctor)}</td>
    <td style="font-size:10px">${getTests(o).map(t => esc(t)).join(', ')}</td>
    <td>${pBadge(o.priority || 'Routine')}</td><td>${sBadge(o.status)}</td>
    <td><button class="ab abO" onclick="openAddResult('${esc(o.id)}')">🔬 Enter</button></td></tr>`).join('') ||
    '<tr><td colspan="7" class="empty"><span class="eIco">✅</span><p>All orders processed!</p></td></tr>';
}

// ═══════════════════════════════════════
// SHIFT
// ═══════════════════════════════════════
function openNewShift() { openM('shiftModal'); sV('smUser', ''); sV('smPass', ''); document.getElementById('smErr').classList.remove('show'); }
async function confirmShift() {
  const u = gV('smUser').trim(), p = gV('smPass');
  const err = document.getElementById('smErr'); err.classList.remove('show');
  const found = ld('users').find(x => x.username === u && x.password === p && x.role === 'lab' && x.status === 'Active');
  if (!found) { err.classList.add('show'); return; }
  const shifts = ld('shifts'), old = ldOne('curShift');
  // BUG FIX: save the old shift into the array BEFORE generating new ID, and don't concat old twice
  if (old && !shifts.find(s => s.id === old.id)) {
    old.endDate = tod(); old.endTime = nowT();
    shifts.push(old);
    await sv('shifts', old); // Save the finished shift
  }
  const nid = genID('SH', shifts);
  const newSh = { id: nid, labUser: found.username, labName: found.name, startDate: tod(), startTime: nowT(), ordersProcessed: 0, resultsEntered: 0, patientsAdded: 0 };
  await svOne('lisV5_curShift', newSh);
  closeM('shiftModal'); toast(`✅ New shift ${nid} started`); renderLabDash();
}
async function incShift(stat, by = 1) { const s = ldOne('curShift'); if (s) { s[stat] = (s[stat] || 0) + by; await svOne('lisV5_curShift', s); } }

// ═══════════════════════════════════════
// USERS CRUD
// ═══════════════════════════════════════
function renderUsers() {
  const q = gV('uSrch').toLowerCase(), rf = gV('uRoleF');
  const filtered = ld('users').filter(u => (!q || (u.username + u.name + u.specialty + u.dept).toLowerCase().includes(q)) && (!rf || u.role === rf));
  document.getElementById('uCount').textContent = filtered.length;
  document.getElementById('usersTbl').innerHTML = filtered.length ? filtered.map(u => `
    <tr><td class="tM">${esc(u.username)}</td><td class="tN">${esc(u.name)}</td><td>${rBadge(u.role)}</td>
    <td style="font-size:11px">${esc(u.specialty || u.labDept || u.dept || '—')}</td>
    <td class="tM">${esc(u.phone || '—')}</td>
    <td><span class="badge ${u.status === 'Active' ? 'bG' : 'bR'}">${esc(u.status)}</span></td>
    <td><button class="ab abE" onclick="editUser('${esc(u.id)}')">✏ Edit</button><button class="ab abD" onclick="delUser('${esc(u.id)}')">🗑</button></td></tr>`).join('')
    : '<tr><td colspan="7" class="empty"><span class="eIco">👥</span><p>No users found</p></td></tr>';
}
function toggleUFields() {
  const r = gV('umRole');
  show('umSpecFd', r === 'doctor'); show('umLicFd', r === 'doctor'); show('umDeptFd', r === 'doctor');
  show('umLabDFd', r === 'lab');
}
function openAddUser(def = 'doctor') {
  document.getElementById('umTitle').textContent = '👥 Add New User';
  document.getElementById('umIdH').value = '';
  ['umName', 'umUser', 'umPass', 'umSpec', 'umLic', 'umDept', 'umPhone', 'umEmail', 'umNotes'].forEach(id => sV(id, ''));
  sV('umRole', def); sV('umStatus', 'Active'); toggleUFields(); openM('userModal');
}
function editUser(id) {
  const u = ld('users').find(x => x.id === id); if (!u) return;
  document.getElementById('umTitle').textContent = '✏ Edit User';
  document.getElementById('umIdH').value = id;
  sV('umRole', u.role); sV('umStatus', u.status); sV('umName', u.name); sV('umUser', u.username); sV('umPass', u.password);
  sV('umSpec', u.specialty || ''); sV('umLic', u.lic || ''); sV('umDept', u.dept || ''); sV('umLabD', u.labDept || '');
  sV('umPhone', u.phone || ''); sV('umEmail', u.email || ''); sV('umNotes', u.notes || '');
  toggleUFields(); openM('userModal');
}
async function saveUser() {
  const name = gV('umName').trim(), uname = gV('umUser').trim().toLowerCase(), pass = gV('umPass').trim();
  if (!name) { toast('Name required', 'error'); return; }
  if (!uname) { toast('Username required', 'error'); return; }
  if (!pass) { toast('Password required', 'error'); return; }
  const users = ld('users'), idH = gV('umIdH'), role = gV('umRole');
  if (users.find(x => x.username === uname && x.id !== idH)) { toast('Username already taken', 'error'); return; }
  const obj = { id: idH || genID('U', users), username: uname, password: pass, role, name, specialty: gV('umSpec'), lic: gV('umLic'), dept: gV('umDept'), labDept: gV('umLabD'), phone: gV('umPhone'), email: gV('umEmail'), notes: gV('umNotes'), status: gV('umStatus'), created: idH ? users.find(x => x.id === idH)?.created || new Date().toISOString() : new Date().toISOString() };
  await sv('users', obj);
  closeM('userModal'); renderUsers(); toast(idH ? '✅ User updated' : '✅ User created');
}
function delUser(id) { conf2('Delete User', 'Delete this user account? Data is preserved.', () => { delRow('users', id); renderUsers(); toast('User deleted'); }); }
function genPass() { const c = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789!@#'; let p = ''; for (let i = 0; i < 10; i++)p += c[Math.floor(Math.random() * c.length)]; sV('umPass', p); }

// ═══════════════════════════════════════
// PATIENTS CRUD
// ═══════════════════════════════════════
function renderPatients() {
  const q = gV('pSrch').toLowerCase(), gf = gV('pGenderF');
  let pats = ld('patients');
  if (CU.role === 'doctor') pats = pats.filter(p => p.ownerId === CU.id);
  const f = pats.filter(p => (!q || (p.name + p.id + p.phone + p.address).toLowerCase().includes(q)) && (!gf || p.gender === gf));
  document.getElementById('pCount').textContent = f.length;
  document.getElementById('patsTbl').innerHTML = f.length ? f.map(p => `
    <tr><td class="tM">${esc(p.id)}</td>
    <td class="tNC" onclick="openProf('${esc(p.id)}')">${esc(p.name)}</td>
    <td class="tM">${p.age ? p.age + 'y' : '—'}</td>
    <td>${p.gender ? `<span class="badge ${p.gender === 'Male' ? 'bB' : 'bP'}">${esc(p.gender)}</span>` : ''}</td>
    <td><span class="badge bGr">${esc(p.blood || '—')}</span></td>
    <td class="tM">${esc(p.phone || '—')}</td>
    <td><span class="badge ${p.type === 'outpatient' ? 'bO' : 'bT'}">${p.type === 'outpatient' ? 'Outpatient' : 'Regular'}</span></td>
    <td style="font-size:10px;color:var(--gray5)">${esc(p.ownerName || 'Admin')}</td>
    <td class="tM">${fmtD(p.date)}</td>
    <td><button class="ab abV" onclick="openProf('${esc(p.id)}')">👁</button><button class="ab abE" onclick="editPat('${esc(p.id)}')">✏</button>${CU.role === 'doctor' ? `<button class="ab" onclick="openPatHistory('${esc(p.id)}')" style="background:var(--purple-lt);color:var(--purple2);border:none;border-radius:5px;padding:4px 9px;font-size:10px;font-weight:700;cursor:pointer">📊 History</button>` : ''} ${(CU.role === 'admin' || (CU.role === 'lab' && p.ownerId === CU.id)) ? `<button class="ab abD" onclick="delPat('${esc(p.id)}')">🗑</button>` : ''}</td></tr>`).join('')
    : '<tr><td colspan="10" class="empty"><span class="eIco">👤</span><p>No patients found</p></td></tr>';
}
function openAddPatient() {
  document.getElementById('pmTitle').textContent = '👤 New Patient'; document.getElementById('pmIdH').value = '';
  ['pmId', 'pmName', 'pmDob', 'pmPhone', 'pmEmail', 'pmAddr', 'pmIns', 'pmNotes', 'pmExtDoctor'].forEach(id => sV(id, ''));
  sV('pmDate', tod()); sV('pmGender', ''); sV('pmBlood', ''); sV('pmAge', ''); sV('pmType', 'regular');
  // Show external doctor field only for lab role
  const extFd = document.getElementById('pmExtDoctorFd');
  if (extFd) { extFd.style.display = (CU.role === 'lab') ? '' : 'none'; }
  sV('pmId', genID('P', ld('patients'))); openM('patModal');
}
function editPat(id) {
  const p = ld('patients').find(x => x.id === id); if (!p) return;
  document.getElementById('pmTitle').textContent = '✏ Edit Patient'; document.getElementById('pmIdH').value = id;
  sV('pmId', p.id); sV('pmDate', p.date); sV('pmName', p.name); sV('pmDob', p.dob); sV('pmAge', p.age);
  sV('pmGender', p.gender); sV('pmBlood', p.blood); sV('pmPhone', p.phone); sV('pmEmail', p.email);
  sV('pmAddr', p.address); sV('pmIns', p.ins); sV('pmType', p.type || 'regular'); sV('pmNotes', p.notes);
  sV('pmExtDoctor', p.extDoctor || '');
  const extFd = document.getElementById('pmExtDoctorFd');
  if (extFd) { extFd.style.display = (CU.role === 'lab') ? '' : 'none'; }
  openM('patModal');
}
async function savePat() {
  const name = gV('pmName').trim(); if (!name) { toast('Name required', 'error'); return; }
  const pats = ld('patients'), idH = gV('pmIdH');
  const ageVal = gV('pmAge');
  const obj = { id: gV('pmId') || genID('P', pats), name, dob: gV('pmDob'), age: ageVal || null, gender: gV('pmGender'), blood: gV('pmBlood'), phone: gV('pmPhone'), email: gV('pmEmail'), address: gV('pmAddr'), ins: gV('pmIns'), type: gV('pmType'), notes: gV('pmNotes'), extDoctor: gV('pmExtDoctor').trim(), date: idH ? pats.find(x => x.id === idH)?.date || new Date().toISOString() : new Date().toISOString(), ownerId: idH ? pats.find(x => x.id === idH)?.ownerId || CU.id : CU.id, ownerName: idH ? pats.find(x => x.id === idH)?.ownerName || CU.name : CU.name, ownerRole: idH ? pats.find(x => x.id === idH)?.ownerRole || CU.role : CU.role };
  if (!idH && CU.role === 'lab') await incShift('patientsAdded');
  await sv('patients', obj);
  closeM('patModal'); renderPatients(); toast(idH ? '✅ Patient updated' : '✅ Patient registered');
}
function delPat(id) { conf2('Delete Patient', 'Permanently delete this patient record?', () => { delRow('patients', id); renderPatients(); toast('Patient deleted'); }); }

// ═══════════════════════════════════════
// PATIENT PROFILE
// ═══════════════════════════════════════
function openProf(pid) {
  const p = ld('patients').find(x => x.id === pid); if (!p) { toast('Patient not found', 'error'); return; }
  const ords = ld('orders').filter(o => o.patientId === pid || o.patient === p.name);
  const res = ld('results').filter(r => r.patientId === pid || r.patient === p.name);
  const bills = ld('billing').filter(b => b.patientId === pid || b.patient === p.name);
  const billed = bills.reduce((s, b) => s + Number(b.total || 0), 0);
  const paid = bills.reduce((s, b) => s + Number(b.paid || 0), 0);
  document.getElementById('profContent').innerHTML = `
    <div class="pHd"><div class="pAv">${p.gender === 'Female' ? '👩' : '👤'}</div>
    <div class="pIn"><h2>${esc(p.name)}</h2><p>${esc(p.id)} · ${esc(p.phone || 'No phone')} · ${esc(p.address || '—')}</p>
    <div class="pMet">${p.gender ? `<span>${esc(p.gender)}</span>` : ''} ${p.age ? `<span>Age ${esc(p.age)}</span>` : ''} ${p.blood ? `<span>Blood: ${esc(p.blood)}</span>` : ''} ${p.ins ? `<span>Ins: ${esc(p.ins)}</span>` : ''}<span>${p.type === 'outpatient' ? 'Outpatient' : 'Regular'}</span><span>Reg: ${fmtD(p.date)}</span></div></div></div>
    <div class="pTabs">
      <div class="pTab active" onclick="swPTab(this,'pti')">📋 Info</div>
      <div class="pTab" onclick="swPTab(this,'pto')">📋 Orders (${ords.length})</div>
      <div class="pTab" onclick="swPTab(this,'ptr')">🔬 Results (${res.length})</div>
      <div class="pTab" onclick="swPTab(this,'ptb')">💰 Billing (${bills.length})</div>
    </div>
    <div class="pTC active" id="pti">
      <div class="pStat">
        <div class="ps"><div class="n" style="color:var(--blue)">${ords.length}</div><div class="l">Orders</div></div>
        <div class="ps"><div class="n" style="color:var(--green2)">${res.length}</div><div class="l">Results</div></div>
        <div class="ps"><div class="n" style="color:var(--orange2)">₪${billed}</div><div class="l">Billed</div></div>
        <div class="ps"><div class="n" style="color:${(billed - paid) > 0 ? 'var(--red2)' : 'var(--green2)'}">₪${billed - paid}</div><div class="l">Balance</div></div>
      </div>
      <div class="iGrid">
        ${[['Full Name', p.name], ['Patient ID', p.id], ['Date of Birth', fmtD(p.dob) || '—'], ['Age', p.age ? p.age + ' years' : '—'], ['Gender', p.gender || '—'], ['Blood Type', p.blood || '—'], ['Phone', p.phone || '—'], ['Insurance', p.ins || '—']].map(([l, v]) => `<div class="iRow"><span class="iLbl">${l}</span><span class="iVal">${esc(v)}</span></div>`).join('')}
        <div class="iRow" style="grid-column:span 2"><span class="iLbl">Address</span><span class="iVal">${esc(p.address || '—')}</span></div>
        <div class="iRow" style="grid-column:span 2"><span class="iLbl">Notes</span><span class="iVal">${esc(p.notes || '—')}</span></div>
        <div class="iRow"><span class="iLbl">Added By</span><span class="iVal">${esc(p.ownerName || '—')}</span></div>
        <div class="iRow"><span class="iLbl">Registered</span><span class="iVal">${fmtD(p.date)}</span></div>
      </div>
    </div>
    <div class="pTC" id="pto">${ords.length ? `<table class="mT"><thead><tr><th>Order ID</th><th>Date</th><th>Doctor</th><th>Tests</th><th>Status</th></tr></thead><tbody>${ords.map(o => `<tr><td class="tM">${esc(o.id)}</td><td class="tM">${fmtD(o.date)}</td><td>${esc(o.doctor)}</td><td style="font-size:10px">${getTests(o).join(', ')}</td><td>${sBadge(o.status)}</td></tr>`).join('')}</tbody></table>` : '<div class="empty"><span class="eIco">📋</span><p>No orders</p></div>'}</div>
    <div class="pTC" id="ptr">${res.length ? `<table class="mT"><thead><tr><th>Test</th><th>Parameter</th><th>Value</th><th>Unit</th><th>Range</th><th>Date</th></tr></thead><tbody>${res.map(r => `<tr><td>${esc(r.testName)}</td><td style="font-size:10px;color:var(--gray5)">${esc(r.paramName || '—')}</td><td class="tM"><b>${esc(r.value)}</b></td><td style="font-size:10px">${esc(r.unit)}</td><td style="font-size:10px">${esc(r.norm)}</td><td class="tM">${fmtD(r.date)}</td></tr>`).join('')}</tbody></table>` : '<div class="empty"><span class="eIco">🔬</span><p>No results yet</p></div>'}</div>
    <div class="pTC" id="ptb">${bills.length ? `<table class="mT"><thead><tr><th>Invoice</th><th>Date</th><th>Total</th><th>Paid</th><th>Balance</th><th>Status</th></tr></thead><tbody>${bills.map(b => `<tr><td class="tM">${esc(b.id)}</td><td class="tM">${fmtD(b.date)}</td><td class="tPr">₪${Number(b.total || 0)}</td><td style="color:var(--green2);font-weight:700">₪${Number(b.paid || 0)}</td><td style="color:${(Number(b.total || 0) - Number(b.paid || 0)) > 0 ? 'var(--red2)' : 'var(--green2)'};font-weight:700">₪${Number(b.total || 0) - Number(b.paid || 0)}</td><td>${Number(b.paid || 0) >= Number(b.total || 0) ? '<span class="badge bG">Paid</span>' : '<span class="badge bR">Unpaid</span>'}</td></tr>`).join('')}</tbody></table>` : '<div class="empty"><span class="eIco">💰</span><p>No billing</p></div>'}</div>`;
  document.getElementById('profEditBtn').onclick = () => { closeM('profModal'); editPat(pid); };
  openM('profModal');
}
function swPTab(el, tid) {
  el.closest('.mXl').querySelectorAll('.pTab').forEach(t => t.classList.remove('active'));
  el.closest('.mXl').querySelectorAll('.pTC').forEach(t => t.classList.remove('active'));
  el.classList.add('active'); document.getElementById(tid).classList.add('active');
}

// ═══════════════════════════════════════
// TESTS CRUD
// ═══════════════════════════════════════
function fillTDeptF() {
  const sel = document.getElementById('tDeptF');
  const depts = [...new Set(ld('tests').map(t => t.dept))].sort();
  sel.innerHTML = '<option value="">All Departments</option>' + depts.map(d => `<option>${esc(d)}</option>`).join('');
}
function renderTests() {
  const q = gV('tSrch').toLowerCase(), df = gV('tDeptF');
  const f = ld('tests').filter(t => (!q || (t.name + t.id + t.dept).toLowerCase().includes(q)) && (!df || t.dept === df));
  document.getElementById('tCount').textContent = f.length;
  document.getElementById('testsTbl').innerHTML = f.length ? f.map(t => `
    <tr><td class="tM">${esc(t.id)}</td><td class="tN">${esc(t.name)}</td>
    <td><span class="badge bT">${esc(t.dept)}</span></td>
    <td style="font-size:10px">${esc(t.sample || '—')}</td>
    <td class="tPr">₪${Number(t.price || 0)}</td>
    <td class="tM">${esc(t.tat || '—')}h</td>
    <td style="font-size:10px">${esc(t.norm || '—')}</td>
    <td><span class="badge ${t.params && t.params.length > 0 ? 'bP' : 'bGr'}">${t.params && t.params.length > 0 ? t.params.length + ' params' : 'Simple'}</span></td>
    <td><span class="badge ${t.status === 'Active' ? 'bG' : 'bR'}">${esc(t.status)}</span></td>
    <td><button class="ab abE" onclick="editTest('${esc(t.id)}')">✏ Edit</button>${(CU.role === 'admin' || CU.role === 'lab') ? `<button class="ab abD" onclick="delTest('${esc(t.id)}')">🗑 Delete</button>` : ''}</td></tr>`).join('')
    : '<tr><td colspan="10" class="empty"><span class="eIco">🧪</span><p>No tests</p></td></tr>';
}
function addParamRow(n = '', u = '', nm = '', nf = '') {
  const c = document.getElementById('paramsBox'), row = document.createElement('div'); row.className = 'pRow';
  row.innerHTML = `<input type="text" placeholder="Parameter name" value="${esc(n)}" class="pN"/><input type="text" placeholder="Unit" value="${esc(u)}" class="pU"/><input type="text" placeholder="Normal (M)" value="${esc(nm)}" class="pNm"/><input type="text" placeholder="Normal (F)" value="${esc(nf)}" class="pNf"/><button type="button" class="rmBtn" onclick="this.closest('.pRow').remove()">✕</button>`;
  c.appendChild(row);
}
function openAddTest() {
  document.getElementById('tmTitle').textContent = '🧪 Add New Test'; document.getElementById('tmIdH').value = '';
  ['tmId', 'tmName', 'tmUnit', 'tmNorm', 'tmNormM', 'tmNormF', 'tmMethod', 'tmNotes', 'tmTat', 'tmPrice'].forEach(id => sV(id, ''));
  sV('tmDept', 'Chemistry'); sV('tmSample', 'Blood (Serum)'); sV('tmStatus', 'Active');
  sV('tmId', genID('T', ld('tests'))); document.getElementById('paramsBox').innerHTML = ''; openM('testModal');
}
function editTest(id) {
  const t = ld('tests').find(x => x.id === id); if (!t) return;
  document.getElementById('tmTitle').textContent = '✏ Edit Test'; document.getElementById('tmIdH').value = id;
  sV('tmId', t.id); sV('tmStatus', t.status); sV('tmName', t.name); sV('tmDept', t.dept); sV('tmSample', t.sample || 'Blood (Serum)');
  sV('tmPrice', t.price); sV('tmTat', t.tat); sV('tmUnit', t.unit); sV('tmMethod', t.method || '');
  sV('tmNorm', t.norm); sV('tmNormM', t.normM || ''); sV('tmNormF', t.normF || ''); sV('tmNotes', t.notes || '');
  document.getElementById('paramsBox').innerHTML = '';
  if (t.params && t.params.length) t.params.forEach(p => addParamRow(p.name, p.unit, p.normM, p.normF)); openM('testModal');
}
async function saveTest() {
  const name = gV('tmName').trim(); if (!name) { toast('Test name required', 'error'); return; }
  const rows = document.getElementById('paramsBox').querySelectorAll('.pRow');
  const params = [...rows].map(r => ({ name: r.querySelector('.pN')?.value.trim() || '', unit: r.querySelector('.pU')?.value.trim() || '', normM: r.querySelector('.pNm')?.value.trim() || '', normF: r.querySelector('.pNf')?.value.trim() || '' })).filter(p => p.name);
  const tests = ld('tests'), idH = gV('tmIdH');
  const obj = { id: gV('tmId') || genID('T', tests), name, status: gV('tmStatus'), dept: gV('tmDept'), sample: gV('tmSample'), price: parseFloat(gV('tmPrice')) || 0, tat: gV('tmTat'), unit: gV('tmUnit'), method: gV('tmMethod'), norm: gV('tmNorm'), normM: gV('tmNormM'), normF: gV('tmNormF'), notes: gV('tmNotes'), params };
  await sv('tests', obj);
  closeM('testModal'); renderTests(); toast(idH ? '✅ Test updated' : '✅ Test added');
}
function delTest(id) { conf2('Delete Test', 'Remove test from catalog?', () => { delRow('tests', id); renderTests(); toast('Test removed'); }); }

// ═══════════════════════════════════════
// ORDERS CRUD
// ═══════════════════════════════════════
function renderOrders() {
  const q = gV('oSrch').toLowerCase(), sf = gV('oStatF'), pf = gV('oPriF');
  let ords = ld('orders');
  if (CU.role === 'doctor') ords = ords.filter(o => o.doctorId === CU.id);
  const f = ords.filter(o => (!q || (o.id + o.patient + o.doctor).toLowerCase().includes(q)) && (!sf || o.status === sf) && (!pf || o.priority === pf)).sort((a, b) => b.id.localeCompare(a.id));
  document.getElementById('ordsTbl').innerHTML = f.length ? f.map(o => `
    <tr><td class="tM">${esc(o.id)}</td><td class="tM">${fmtD(o.date)}</td><td class="tN">${esc(o.patient)}</td>
    <td style="font-size:11px">${esc(o.doctor)}</td>
    <td style="font-size:10px;max-width:150px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${esc(getTests(o).join(', '))}">${getTests(o).map(t => `<span class="badge bGr" style="margin:1px">${esc(t)}</span>`).join('') || '—'}</td>
    <td>${pBadge(o.priority || 'Routine')}</td><td>${sBadge(o.status)}</td>
    <td>${CU.role === 'lab' ? `<button class="ab abO" onclick="openAddResult('${esc(o.id)}')">🔬 Enter</button>` : ''}
    ${CU.role === 'doctor' ? `<button class="ab" onclick="viewOrderResults('${esc(o.id)}')" style="background:var(--teal-lt);color:var(--teal);border:none;border-radius:5px;padding:4px 9px;font-size:10px;font-weight:700;cursor:pointer">🔬 Results</button>` : ''}
    ${(CU.role === 'admin' || CU.role === 'lab') ? `<button class="ab" onclick="viewOrderResults('${esc(o.id)}')" style="background:var(--sky-lt);color:var(--blue2);border:none;border-radius:5px;padding:4px 9px;font-size:10px;font-weight:700;cursor:pointer">👁 View</button>` : ''}
    <button class="ab abE" onclick="editOrder('${esc(o.id)}')">✏</button>
    <button class="ab abD" onclick="delOrder('${esc(o.id)}')">🗑</button></td></tr>`).join('')
    : '<tr><td colspan="8" class="empty"><span class="eIco">📋</span><p>No orders found</p></td></tr>';
}
function addTestRow() {
  const c = document.getElementById('ordTestsList'), tests = ld('tests').filter(t => t.status === 'Active'), i = c.querySelectorAll('.oTR').length + 1;
  const row = document.createElement('div'); row.className = 'oTR';
  row.innerHTML = `<span class="tNum">${i}</span><select style="flex:1;border:1.5px solid var(--gray2);border-radius:5px;padding:6px 8px;font-size:12px;font-family:inherit;outline:none"><option value="">Select test...</option>${tests.map(t => `<option value="${esc(t.name)}">${esc(t.name)} (₪${t.price})</option>`).join('')}</select><button type="button" class="rmBtn" onclick="this.closest('.oTR').remove()">✕</button>`;
  c.appendChild(row);
}
function populateOrdModal() {
  let pats = ld('patients'); if (CU.role === 'doctor') pats = pats.filter(p => p.ownerId === CU.id);
  document.getElementById('omPat').innerHTML = '<option value="">Select patient...</option>' + pats.map(p => `<option value="${esc(p.name)}" data-pid="${esc(p.id)}">${esc(p.name)} (${esc(p.id)})</option>`).join('');
  const docSel = document.getElementById('omDoc');
  const extFd = document.getElementById('omExtDocFd');
  if (CU.role === 'doctor') {
    docSel.innerHTML = `<option value="${esc(CU.name)}">${esc(CU.name)}</option>`; docSel.disabled = true;
    if (extFd) extFd.classList.add('hn');
  } else {
    const docs = ld('users').filter(u => u.role === 'doctor' && u.status === 'Active');
    docSel.innerHTML = '<option value="">Select system doctor...</option>' + docs.map(d => `<option value="${esc(d.name)}">${esc(d.name)}</option>`).join('');
    docSel.innerHTML += `<option value="__none__">— No system doctor —</option>`;
    docSel.disabled = false;
    if (extFd) extFd.classList.remove('hn');
  }
}
function onOrdPatChange() {
  // When lab selects a patient who has an extDoctor, pre-fill the external doctor field
  if (CU.role !== 'lab') return;
  const patName = gV('omPat');
  const p = ld('patients').find(x => x.name === patName);
  if (p && p.extDoctor) { sV('omExtDoc', p.extDoctor); }
}
function openAddOrder() {
  document.getElementById('omTitle').textContent = '📋 New Lab Order'; document.getElementById('omIdH').value = '';
  sV('omId', genID('ORD', ld('orders'))); sV('omDate', tod()); sV('omStat', 'Pending'); sV('omPri', 'Routine'); sV('omNotes', ''); sV('omExtDoc', '');
  populateOrdModal(); document.getElementById('ordTestsList').innerHTML = ''; addTestRow(); openM('ordModal');
}
function editOrder(id) {
  const o = ld('orders').find(x => x.id === id); if (!o) return;
  document.getElementById('omTitle').textContent = '✏ Edit Order'; document.getElementById('omIdH').value = id;
  sV('omId', o.id); sV('omDate', o.date); sV('omStat', o.status); sV('omPri', o.priority || 'Routine'); sV('omNotes', o.notes);
  populateOrdModal(); sV('omPat', o.patient); sV('omDoc', o.doctor || ''); sV('omExtDoc', o.extDoctor || '');
  const tests = ld('tests').filter(t => t.status === 'Active'), c = document.getElementById('ordTestsList'); c.innerHTML = '';
  getTests(o).forEach((tn, i) => { const row = document.createElement('div'); row.className = 'oTR'; row.innerHTML = `<span class="tNum">${i + 1}</span><select style="flex:1;border:1.5px solid var(--gray2);border-radius:5px;padding:6px 8px;font-size:12px;font-family:inherit;outline:none"><option value="">Select...</option>${tests.map(t => `<option value="${esc(t.name)}" ${t.name === tn ? 'selected' : ''}>${esc(t.name)} (₪${t.price})</option>`).join('')}</select><button type="button" class="rmBtn" onclick="this.closest('.oTR').remove()">✕</button>`; c.appendChild(row); });
  if (!getTests(o).length) addTestRow(); openM('ordModal');
}
async function saveOrder() {
  const pat = gV('omPat');
  if (!pat) { toast('Select a patient', 'error'); return; }
  // Determine doctor: system dropdown or external free-text (lab only)
  let docName = gV('omDoc');
  const extDocRaw = (gV('omExtDoc') || '').trim();
  if (CU.role === 'lab') {
    // If dropdown is blank or "__none__", use external doctor name
    if (!docName || docName === '__none__') {
      if (!extDocRaw) { toast('Enter a doctor name in the dropdown or the external doctor field', 'error'); return; }
      docName = extDocRaw;
    }
  } else {
    if (!docName) { toast('Select a doctor', 'error'); return; }
  }
  const rows = document.getElementById('ordTestsList').querySelectorAll('.oTR select');
  const tests = [...rows].map(s => s.value).filter(Boolean);
  if (!tests.length) { toast('Add at least one test', 'error'); return; }
  const ords = ld('orders'), idH = gV('omIdH');
  const docU = ld('users').find(u => u.name === docName);
  const patO = ld('patients').find(p => p.name === pat);
  const sh = ldOne('curShift');
  // Bug fix: lab users creating orders with external (non-system) doctors should NOT have their own id stored as doctorId
  const resolvedDocId = docU ? docU.id : (CU.role === 'doctor') ? CU.id : '';
  const obj = { id: gV('omId') || genID('ORD', ords), date: idH ? ords.find(x => x.id === idH)?.date || new Date().toISOString() : new Date().toISOString(), patient: pat, patientId: patO?.id || '', doctorId: resolvedDocId, doctor: docName, extDoctor: extDocRaw || '', tests, status: gV('omStat'), priority: gV('omPri') || 'Routine', notes: gV('omNotes'), shiftId: sh?.id || '' };
  if (!idH && CU.role === 'lab') await incShift('ordersProcessed');
  await sv('orders', obj);
  closeM('ordModal'); renderOrders(); toast(idH ? '✅ Order updated' : '✅ Order created');
}
function delOrder(id) { conf2('Delete Order', 'Delete this lab order?', () => { delRow('orders', id); renderOrders(); toast('Order deleted'); }); }

// ═══════════════════════════════════════
// RESULTS CRUD
// ═══════════════════════════════════════
function renderResults() {
  const q = gV('rSrch').toLowerCase();
  const f = ld('results').filter(r => (!q || (r.testName + r.paramName + r.patient + r.orderId).toLowerCase().includes(q))).sort((a, b) => b.id.localeCompare(a.id));
  document.getElementById('resSumm').textContent = `${f.length} result${f.length !== 1 ? 's' : ''}`;
  // Bug fix: cache patients lookup outside the map loop to avoid repeated localStorage deserializations
  const _allPats = ld('patients');
  document.getElementById('resTbl').innerHTML = f.length ? f.map(r => `
    <tr><td class="tM">${esc(r.id)}</td><td class="tM">${esc(r.orderId || '—')}</td>
    <td style="font-size:11px;font-weight:600">${esc(r.testName)}</td>
    <td style="font-size:10px;color:var(--gray5)">${esc(r.paramName || '—')}</td>
    <td class="tNC" onclick="openProf('${esc(_allPats.find(p => p.name === r.patient || p.id === r.patientId)?.id || '')}')">${esc(r.patient)}</td>
    <td class="tM"><b>${esc(r.value)}</b></td><td style="font-size:10px">${esc(r.unit || '—')}</td>
    <td style="font-size:10px">${esc(r.norm || '—')}</td>
    <td style="font-size:11px">${esc(r.tech || '—')}</td><td class="tM">${fmtD(r.date)}</td>
    <td>${CU.role === 'lab' ? `<button class="ab abD" onclick="delResult('${esc(r.id)}')">🗑</button>` : ''}</td></tr>`).join('')
    : '<tr><td colspan="11" class="empty"><span class="eIco">🔬</span><p>No results</p></td></tr>';
}
function openAddResult(preOid) {
  if (CU.role !== 'lab') { toast('Only lab technicians can enter results', 'error'); return; }
  document.getElementById('rmTitle').textContent = '🔬 Enter Lab Results';
  const ords = ld('orders').filter(o => o.status !== 'Cancelled');
  const sel = document.getElementById('rmOrd');
  sel.innerHTML = '<option value="">— Select Order —</option>' + ords.map(o => `<option value="${esc(o.id)}" ${o.id === preOid ? 'selected' : ''}>${esc(o.id)} — ${esc(o.patient)} (${getTests(o).length} tests)</option>`).join('');
  sV('rmDate', tod()); sV('rmTech', CU.username); sV('rmPat', ''); sV('rmDoc', '');
  document.getElementById('resTestsCont').innerHTML = '<div style="text-align:center;color:var(--gray4);padding:28px;border:2px dashed var(--gray2);border-radius:8px;font-size:13px">☝️ Select an order to load tests</div>';
  openM('resModal'); if (preOid) { sV('rmOrd', preOid); loadOrdForResult(); }
}
function loadOrdForResult() {
  const oid = gV('rmOrd'); if (!oid) { document.getElementById('resTestsCont').innerHTML = '<div style="text-align:center;color:var(--gray4);padding:28px;border:2px dashed var(--gray2);border-radius:8px;font-size:13px">☝️ Select an order to load tests</div>'; return; }
  const ord = ld('orders').find(o => o.id === oid); if (!ord) return;
  sV('rmPat', ord.patient); sV('rmDoc', ord.doctor);
  const tests = ld('tests'), existing = ld('results').filter(r => r.orderId === oid), ordTests = getTests(ord);
  // Lab account: flag column is always hidden — lab doesn't manually set flags
  let html = `<table class="rET"><thead><tr><th>Test / Parameter</th><th>Value *</th><th>Unit</th><th>Normal Range</th><th>Comment</th></tr></thead><tbody>`;
  ordTests.forEach(tn => {
    const tObj = tests.find(t => t.name === tn);
    if (tObj && tObj.params && tObj.params.length > 0) {
      html += `<tr style="background:var(--navy2)"><td colspan="5" style="font-size:11px;font-weight:700;color:#90CAF9;padding:7px 10px">${esc(tn)}</td></tr>`;
      tObj.params.forEach(pr => {
        const ex = existing.find(r => r.testName === tn && r.paramName === pr.name);
        html += `<tr data-t="${esc(tn)}" data-p="${esc(pr.name)}"><td style="padding-left:20px;font-size:11px;color:var(--gray6)">↳ ${esc(pr.name)}</td><td><input class="rI" name="val" type="text" value="${esc(ex?.value || '')}" placeholder="Enter value"/></td><td style="font-size:10px">${esc(pr.unit)}</td><td style="font-size:10px">${esc(pr.normM || pr.normF || tObj.norm || '—')}</td><td><input class="rI" name="comm" type="text" value="${esc(ex?.comment || '')}" placeholder="Comment"/></td></tr>`;
      });
    } else {
      const ex = existing.find(r => r.testName === tn && !r.paramName);
      html += `<tr data-t="${esc(tn)}" data-p=""><td style="font-size:11px;font-weight:600">${esc(tn)}</td><td><input class="rI" name="val" type="text" value="${esc(ex?.value || '')}" placeholder="Result value"/></td><td style="font-size:10px">${esc(tObj?.unit || '—')}</td><td style="font-size:10px">${esc(tObj?.norm || '—')}</td><td><input class="rI" name="comm" type="text" value="${esc(ex?.comment || '')}" placeholder="Comment"/></td></tr>`;
    }
  });
  html += '</tbody></table>';
  document.getElementById('resTestsCont').innerHTML = html;
}
async function saveResults() {
  const oid = gV('rmOrd'); if (!oid) { toast('Select an order', 'error'); return; }
  const ord = ld('orders').find(o => o.id === oid); if (!ord) return;
  const rows = document.getElementById('resTestsCont').querySelectorAll('tbody tr[data-t]');
  if (!rows.length) { toast('No test rows', 'error'); return; }
  const tests = ld('tests');
  // Bug fix: keep a reference to ALL results for correct ID generation,
  // then build the filtered working array separately to avoid ID collisions
  // when re-entering results for a previously-completed order.
  const allResults = ld('results');
  let count = 0;
  for (const row of rows) {
    const tn = row.getAttribute('data-t'), pn = row.getAttribute('data-p') || '';
    const val = row.querySelector('[name="val"]')?.value.trim(); if (!val) continue;
    const comm = row.querySelector('[name="comm"]')?.value || '';
    const tObj = tests.find(t => t.name === tn);
    let unit = '', norm = '';
    if (tObj) { if (pn) { const p = tObj.params?.find(x => x.name === pn); unit = p?.unit || ''; norm = p?.normM || p?.normF || tObj.norm || ''; } else { unit = tObj.unit || ''; norm = tObj.norm || ''; } }
    // Use a unique ID for results. Combining R + timestamp + random to avoid collisions.
    const newId = 'R' + Date.now() + Math.round(Math.random() * 1000);
    const resultObj = { id: newId, date: new Date().toISOString(), orderId: oid, patient: ord.patient, patientId: ord.patientId || '', testId: tObj?.id || '', testName: tn, paramName: pn, value: val, unit, norm, tech: gV('rmTech') || CU.username, comment: comm };
    await sv('results', resultObj);
    count++;
  }
  if (!count) { toast('Enter at least one value', 'warn'); return; }
  const ords = ld('orders'), oi = ords.findIndex(o => o.id === oid);
  if (oi >= 0 && ords[oi].status !== 'Cancelled') {
    ords[oi].status = 'Completed';
    await sv('orders', ords[oi]);
  }
  if (CU.role === 'lab') await incShift('resultsEntered', count);
  closeM('resModal');
  if (document.getElementById('results-view').classList.contains('active')) renderResults();
  if (document.getElementById('labDash-view').classList.contains('active')) renderLabDash();
  toast(`✅ ${count} results saved — order marked Completed`);
}
function delResult(id) { if (CU.role !== 'lab') { toast('Only lab technicians can delete results', 'error'); return; } conf2('Delete Result', 'Delete this result record?', () => { delRow('results', id); renderResults(); toast('Result deleted'); }); }

// ═══════════════════════════════════════
// VIEW ORDER RESULTS (Doctor)
// ═══════════════════════════════════════
function viewOrderResults(oid) {
  const ord = ld('orders').find(o => o.id === oid);
  if (!ord) { toast('Order not found', 'error'); return; }
  const res = ld('results').filter(r => r.orderId === oid);
  document.getElementById('ordResTitle').textContent = `🔬 Results — ${oid} · ${esc(ord.patient)}`;
  if (!res.length) {
    document.getElementById('ordResContent').innerHTML = `<div style="text-align:center;padding:36px;color:var(--gray4);font-size:13px"><span style="font-size:32px;display:block;margin-bottom:10px">🔬</span>No results entered for this order yet.<br><span style="font-size:11px">Waiting for lab technician to enter results.</span></div>`;
  } else {
    const byTest = {}; res.forEach(r => { const k = r.testName || '—'; if (!byTest[k]) byTest[k] = []; byTest[k].push(r); });
    let html = `<div style="background:var(--sky-lt);border:1px solid var(--blue);border-radius:8px;padding:10px 14px;margin-bottom:12px;font-size:11px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px">
      <div><span style="color:var(--gray5);font-size:9px;font-weight:700;text-transform:uppercase">Order</span><div style="font-weight:700">${esc(oid)}</div></div>
      <div><span style="color:var(--gray5);font-size:9px;font-weight:700;text-transform:uppercase">Patient</span><div style="font-weight:700">${esc(ord.patient)}</div></div>
      <div><span style="color:var(--gray5);font-size:9px;font-weight:700;text-transform:uppercase">Date</span><div style="font-weight:700">${fmtD(res[0]?.date || ord.date)}</div></div>
    </div>`;
    Object.entries(byTest).forEach(([tname, rows]) => {
      html += `<div style="margin-bottom:12px"><div style="font-size:11px;font-weight:800;color:var(--teal);background:var(--teal-lt);padding:6px 10px;border-radius:6px 6px 0 0;border:1px solid var(--teal)">${esc(tname)}</div>
      <table style="width:100%;border-collapse:collapse;border:1px solid var(--gray2);border-top:none">
      <thead><tr style="background:var(--navy2)"><th style="padding:7px 10px;font-size:9px;font-weight:700;color:#90CAF9;text-align:left">Parameter</th><th style="padding:7px 10px;font-size:9px;font-weight:700;color:#90CAF9;text-align:left">Result</th><th style="padding:7px 10px;font-size:9px;font-weight:700;color:#90CAF9;text-align:left">Unit</th><th style="padding:7px 10px;font-size:9px;font-weight:700;color:#90CAF9;text-align:left">Normal Range</th><th style="padding:7px 10px;font-size:9px;font-weight:700;color:#90CAF9;text-align:left">Comment</th></tr></thead><tbody>`;
      rows.forEach(r => {
        html += `<tr style="border-bottom:1px solid var(--gray1)">
          <td style="padding:7px 10px;font-size:11px;color:var(--gray6)">${esc(r.paramName || r.testName)}</td>
          <td style="padding:7px 10px;font-size:13px;font-weight:800;font-family:monospace">${esc(r.value)}</td>
          <td style="padding:7px 10px;font-size:10px">${esc(r.unit || '—')}</td>
          <td style="padding:7px 10px;font-size:10px">${esc(r.norm || '—')}</td>
          <td style="padding:7px 10px;font-size:10px;color:var(--gray5)">${esc(r.comment || '—')}</td>
        </tr>`;
      });
      html += `</tbody></table></div>`;
    });
    document.getElementById('ordResContent').innerHTML = html;
  }
  openM('ordResModal');
}

// ═══════════════════════════════════════
// PATIENT HISTORY TABLE (Doctor)
// ═══════════════════════════════════════
function getHistColsKey() { return `lisV5_histCols_${CU.id}`; }
function getHistDataKey(pid) { return `lisV5_histData_${CU.id}_${pid}`; }
function getHistCols() { return DB[getHistColsKey()] || []; }
function getHistData(pid) { return DB[getHistDataKey(pid)] || []; }
async function saveHistData(pid, d) { await svOne(getHistDataKey(pid), d); }

const HIST_TYPES = ['Text', 'Number', 'Date', 'Yes/No', 'Score (1-5)'];

let _histOpenPid = null;
function openPatHistory(pid) {
  const p = ld('patients').find(x => x.id === pid); if (!p) return;
  _histOpenPid = pid;
  document.getElementById('histModalTitle').textContent = `📊 History — ${p.name}`;
  renderHistTable(pid);
  openM('histModal');
}
function renderHistTable(pid) {
  const cols = getHistCols();
  const rows = getHistData(pid);
  let html = '';
  if (!cols.length) {
    html = `<div style="text-align:center;padding:32px;color:var(--gray4)"><span style="font-size:32px;display:block;margin-bottom:8px">📊</span><p style="font-size:12px">No history columns defined yet.<br>Click <b>⚙ Columns</b> to add your custom columns.</p></div>`;
  } else {
    // Editable new row + existing rows
    html += `<div style="margin-bottom:14px">
      <div style="font-size:11px;font-weight:700;color:var(--gray6);margin-bottom:6px;text-transform:uppercase;letter-spacing:.4px">➕ Add New Entry</div>
      <div style="display:flex;gap:8px;align-items:flex-end;flex-wrap:wrap;background:var(--teal-lt);padding:10px;border-radius:8px;border:1px solid var(--teal)">
        ${cols.map(c => `<div style="display:flex;flex-direction:column;gap:3px;min-width:110px;flex:1">
          <label style="font-size:9px;font-weight:700;color:var(--teal);text-transform:uppercase">${esc(c.name)}</label>
          ${histInputForType(c, 'hist_new_' + sanitizeId(c.name))}
        </div>`).join('')}
      </div>
    </div>`;
    html += `<div style="font-size:11px;font-weight:700;color:var(--gray6);margin-bottom:6px;text-transform:uppercase;letter-spacing:.4px">📋 History (${rows.length} entries)</div>`;
    if (!rows.length) {
      html += '<div style="text-align:center;padding:20px;color:var(--gray4);font-size:12px">No entries yet.</div>';
    } else {
      html += `<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;min-width:400px">
        <thead><tr style="background:var(--navy2)">
          <th style="padding:7px 10px;font-size:9px;font-weight:700;color:#90CAF9;text-align:left">Date</th>
          ${cols.map(c => `<th style="padding:7px 10px;font-size:9px;font-weight:700;color:#90CAF9;text-align:left">${esc(c.name)}</th>`).join('')}
          <th style="padding:7px 10px;font-size:9px;font-weight:700;color:#90CAF9;text-align:center">Del</th>
        </tr></thead><tbody>`;
      rows.slice().reverse().forEach((row, ri) => {
        const realIdx = rows.length - 1 - ri;
        html += `<tr style="border-bottom:1px solid var(--gray1)">
          <td style="padding:6px 10px;font-size:10px;font-family:monospace;color:var(--gray5)">${esc(row._date || '—')}</td>
          ${cols.map(c => `<td style="padding:6px 10px;font-size:12px">${esc(row[c.name] != null ? row[c.name] : '—')}</td>`).join('')}
          <td style="padding:6px 10px;text-align:center"><button class="ab abD" style="padding:2px 7px;font-size:10px" onclick="delHistRow('${esc(pid)}',${realIdx})">🗑</button></td>
        </tr>`;
      });
      html += `</tbody></table></div>`;
    }
  }
  document.getElementById('histContent').innerHTML = html;
}
function histInputForType(col, id) {
  const t = col.type || 'Text';
  if (t === 'Date') return `<input type="date" id="${id}" style="border:1.5px solid var(--gray2);border-radius:6px;padding:6px 8px;font-size:12px;font-family:inherit;width:100%;outline:none"/>`;
  if (t === 'Number') return `<input type="number" id="${id}" placeholder="0" style="border:1.5px solid var(--gray2);border-radius:6px;padding:6px 8px;font-size:12px;font-family:inherit;width:100%;outline:none"/>`;
  if (t === 'Yes/No') return `<select id="${id}" style="border:1.5px solid var(--gray2);border-radius:6px;padding:6px 8px;font-size:12px;font-family:inherit;width:100%;outline:none"><option value="">—</option><option>Yes</option><option>No</option></select>`;
  if (t === 'Score (1-5)') return `<select id="${id}" style="border:1.5px solid var(--gray2);border-radius:6px;padding:6px 8px;font-size:12px;font-family:inherit;width:100%;outline:none"><option value="">—</option><option>1</option><option>2</option><option>3</option><option>4</option><option>5</option></select>`;
  return `<input type="text" id="${id}" placeholder="Enter value" style="border:1.5px solid var(--gray2);border-radius:6px;padding:6px 8px;font-size:12px;font-family:inherit;width:100%;outline:none"/>`;
}
function sanitizeId(s) { return s.replace(/[^a-zA-Z0-9]/g, '_'); }
async function saveHistRow() {
  if (!_histOpenPid) { return; }
  const cols = getHistCols(); if (!cols.length) { toast('Add columns first via ⚙ Columns', 'warn'); return; }
  const row = { _date: new Date().toISOString() }; let hasVal = false;
  cols.forEach(c => {
    const el = document.getElementById('hist_new_' + sanitizeId(c.name));
    if (el) { const v = el.value.trim(); if (v) hasVal = true; row[c.name] = v; }
  });
  if (!hasVal) { toast('Enter at least one value', 'error'); return; }
  const rows = getHistData(_histOpenPid); rows.push(row);
  await saveHistData(_histOpenPid, rows); renderHistTable(_histOpenPid); toast('✅ History entry saved');
}
function delHistRow(pid, idx) {
  const rows = getHistData(pid); rows.splice(idx, 1); saveHistData(pid, rows); renderHistTable(pid); toast('Entry deleted');
}
function openManageHistCols() {
  closeM('histModal');
  const cols = getHistCols();
  document.getElementById('histColsBox').innerHTML = '';
  if (cols.length) { cols.forEach(c => addHistColRow(c.name, c.type)); } else { addHistColRow(); }
  openM('histColsModal');
}
function addHistColRow(name = '', type = 'Text') {
  const box = document.getElementById('histColsBox'), row = document.createElement('div');
  row.className = 'histCR'; row.style.cssText = 'display:grid;grid-template-columns:2fr 1fr 28px;gap:6px;margin-bottom:6px;align-items:end';
  row.innerHTML = `<input type="text" value="${esc(name)}" placeholder="Column name e.g. Blood Pressure" style="border:1.5px solid var(--gray2);border-radius:6px;padding:7px 9px;font-size:12px;font-family:inherit;outline:none;width:100%"/>
    <select style="border:1.5px solid var(--gray2);border-radius:6px;padding:7px 8px;font-size:12px;font-family:inherit;outline:none;width:100%;cursor:pointer">${HIST_TYPES.map(t => `<option ${t === type ? 'selected' : ''}>${esc(t)}</option>`).join('')}</select>
    <button type="button" style="background:var(--red-lt);color:var(--red2);border:none;border-radius:4px;width:28px;height:34px;cursor:pointer;font-size:13px;display:flex;align-items:center;justify-content:center" onclick="this.closest('.histCR').remove()">✕</button>`;
  box.appendChild(row);
}
async function saveHistCols() {
  const rows = document.getElementById('histColsBox').querySelectorAll('.histCR');
  const cols = [...rows].map(r => { const inputs = r.querySelectorAll('input,select'); return { name: (inputs[0]?.value || '').trim(), type: inputs[1]?.value || 'Text' }; }).filter(c => c.name);
  if (!cols.length) { toast('Add at least one column', 'error'); return; }
  const names = cols.map(c => c.name); if (new Set(names).size !== names.length) { toast('Column names must be unique', 'error'); return; }
  await svOne(getHistColsKey(), cols);
  closeM('histColsModal');
  if (_histOpenPid) { openM('histModal'); renderHistTable(_histOpenPid); }
  toast('✅ Columns saved');
}

// ═══════════════════════════════════════
// AUTH MODAL HELPER
// ═══════════════════════════════════════
function promptAuth(title, desc, cb) {
  pendingAuthCb = cb;
  document.getElementById('authTitle').textContent = title;
  document.getElementById('authDesc').textContent = desc;
  document.getElementById('authErr').classList.remove('show');
  sV('authUser', CU.username); sV('authPass', '');
  openM('authModal');
}
function confirmAuthModal() {
  const u = gV('authUser').trim(), p = gV('authPass');
  const err = document.getElementById('authErr'); err.classList.remove('show');
  let ok = false;
  if (CU.role === 'admin') { ok = (u === ADMIN.u && p === ADMIN.p); }
  else { const found = ld('users').find(x => x.username === u && x.password === p && x.id === CU.id); ok = !!found; }
  if (!ok) { err.classList.add('show'); sV('authPass', ''); return; }
  closeM('authModal'); if (pendingAuthCb) { const cb = pendingAuthCb; pendingAuthCb = null; cb(); }
}
// ═══════════════════════════════════════
// BILLING CRUD
// ═══════════════════════════════════════
function renderBilling() {
  const q = gV('bSrch').toLowerCase();
  let bills = ld('billing'); if (CU.role === 'doctor') bills = bills.filter(b => b.doctorId === CU.id);
  const f = bills.filter(b => !q || (b.id + b.patient + b.doctor).toLowerCase().includes(q)).sort((a, b) => b.id.localeCompare(a.id));
  const tR = f.reduce((s, b) => s + Number(b.total || 0), 0), tP = f.reduce((s, b) => s + Number(b.paid || 0), 0);
  document.getElementById('bTotR').textContent = '₪' + tR.toLocaleString();
  document.getElementById('bTotP').textContent = '₪' + tP.toLocaleString();
  document.getElementById('bTotB').textContent = '₪' + (tR - tP).toLocaleString();
  document.getElementById('bInvC').textContent = f.length;
  document.getElementById('billTbl').innerHTML = f.length ? f.map(b => {
    const bal = Number(b.total || 0) - Number(b.paid || 0);
    return `<tr><td class="tM">${esc(b.id)}</td><td class="tN">${esc(b.patient)}</td><td style="font-size:11px">${esc(b.doctor || '—')}</td><td class="tM">${fmtD(b.date)}</td><td class="tPr">₪${Number(b.total || 0)}</td><td style="color:var(--green2);font-weight:700">₪${Number(b.paid || 0)}</td><td style="color:${bal > 0 ? 'var(--red2)' : 'var(--green2)'};font-weight:700">₪${bal}</td><td><span class="badge ${bal <= 0 ? 'bG' : Number(b.paid || 0) > 0 ? 'bGd' : 'bR'}">${bal <= 0 ? 'Paid' : Number(b.paid || 0) > 0 ? 'Partial' : 'Unpaid'}</span></td><td><button class="ab abE" onclick="editBill('${esc(b.id)}')">✏</button><button class="ab abD" onclick="delBill('${esc(b.id)}')">🗑</button></td></tr>`;
  }).join('')
    : '<tr><td colspan="9" class="empty"><span class="eIco">💰</span><p>No invoices</p></td></tr>';
}
function populateBillModal() {
  let pats = ld('patients'); if (CU.role === 'doctor') pats = pats.filter(p => p.ownerId === CU.id);
  document.getElementById('bmPat').innerHTML = '<option value="">Select patient...</option>' + pats.map(p => `<option value="${esc(p.name)}">${esc(p.name)}</option>`).join('');
  const docs = ld('users').filter(u => u.role === 'doctor' && u.status === 'Active');
  const ds = document.getElementById('bmDoc'); ds.innerHTML = '<option value="">Select doctor...</option>' + docs.map(d => `<option value="${esc(d.name)}">${esc(d.name)}</option>`).join('');
  if (CU.role === 'doctor') { sV('bmDoc', CU.name); ds.disabled = true; } else ds.disabled = false;
}
function openAddBilling() {
  document.getElementById('bmTitle').textContent = '💰 New Invoice'; document.getElementById('bmIdH').value = '';
  sV('bmId', genID('INV', ld('billing'))); sV('bmDate', tod()); sV('bmTotal', ''); sV('bmPaid', ''); sV('bmNotes', '');
  populateBillModal(); openM('billModal');
}
function editBill(id) {
  const b = ld('billing').find(x => x.id === id); if (!b) return;
  document.getElementById('bmTitle').textContent = '✏ Edit Invoice'; document.getElementById('bmIdH').value = id;
  sV('bmId', b.id); sV('bmDate', b.date); sV('bmTotal', b.total); sV('bmPaid', b.paid); sV('bmNotes', b.notes);
  populateBillModal(); sV('bmPat', b.patient); sV('bmDoc', b.doctor); openM('billModal');
}
async function saveBill() {
  const pat = gV('bmPat'); if (!pat) { toast('Select a patient', 'error'); return; }
  const bills = ld('billing'), idH = gV('bmIdH'), doc = gV('bmDoc');
  const docU = ld('users').find(u => u.name === doc), patO = ld('patients').find(p => p.name === pat);
  // Bug fix: same as saveOrder — lab should not be stored as doctor
  const resolvedDocId2 = docU ? docU.id : (CU.role === 'doctor') ? CU.id : '';
  const obj = { id: gV('bmId') || genID('INV', bills), date: idH ? bills.find(x => x.id === idH)?.date || new Date().toISOString() : new Date().toISOString(), patient: pat, patientId: patO?.id || '', doctorId: resolvedDocId2, doctor: doc, total: parseFloat(gV('bmTotal')) || 0, paid: parseFloat(gV('bmPaid')) || 0, notes: gV('bmNotes') };
  await sv('billing', obj);
  closeM('billModal'); renderBilling(); toast(idH ? '✅ Invoice updated' : '✅ Invoice created');
}
function delBill(id) { conf2('Delete Invoice', 'Delete this invoice?', () => { delRow('billing', id); renderBilling(); toast('Invoice deleted'); }); }

// ═══════════════════════════════════════
// REPORT
// ═══════════════════════════════════════
function populateRepSel() {
  let pats = ld('patients');
  if (CU.role === 'doctor') pats = pats.filter(p => p.ownerId === CU.id);
  // Lab and admin can print reports for all patients
  document.getElementById('rptPatSel').innerHTML = '<option value="">— Select Patient —</option>' + pats.map(p => `<option value="${esc(p.id)}">${esc(p.name)} (${esc(p.id)})</option>`).join('');
  document.getElementById('rptOrdSel').innerHTML = '<option value="">— Select Order —</option>';
  document.getElementById('rptDate').textContent = 'Date: ' + new Date().toLocaleDateString('en-GB');
  document.getElementById('rptBody').innerHTML = '<tr><td colspan="6" style="text-align:center;color:#94A3B8;padding:16px">Select patient and order above</td></tr>';
}
function loadRepOrders() {
  const pid = gV('rptPatSel'); if (!pid) return;
  const p = ld('patients').find(x => x.id === pid);
  const ords = ld('orders').filter(o => o.patientId === pid || o.patient === (p?.name || ''));
  document.getElementById('rptOrdSel').innerHTML = '<option value="">— Select Order —</option>' + ords.map(o => `<option value="${esc(o.id)}">${esc(o.id)} (${fmtD(o.date)})</option>`).join('');
  document.getElementById('rptName').textContent = p?.name || '—';
  document.getElementById('rptPID').textContent = pid;
  document.getElementById('rptGA').textContent = (p?.gender || '—') + ' / Age ' + (p?.age || '—');
}
function loadRepResults() {
  const oid = gV('rptOrdSel'); if (!oid) return;
  const ord = ld('orders').find(o => o.id === oid); if (!ord) return;
  // Show referred-by doctor: prefer order doctor, fallback to extDoctor, fallback to patient extDoctor
  const pat = ld('patients').find(p => p.id === gV('rptPatSel'));
  const docDisplay = ord.doctor && ord.doctor !== '__none__' ? ord.doctor : (ord.extDoctor || pat?.extDoctor || '—');
  document.getElementById('rptDoctor').textContent = docDisplay;
  document.getElementById('rptOID').textContent = oid;
  document.getElementById('rptTDate').textContent = fmtD(ord.date);
  const res = ld('results').filter(r => r.orderId === oid);
  if (!res.length) {
    document.getElementById('rptBody').innerHTML = '<tr><td colspan="5" style="text-align:center;padding:16px;color:#94A3B8">No results entered for this order</td></tr>';
    return;
  }
  // Group results by test name
  const byTest = {}; res.forEach(r => { const k = r.testName || '—'; if (!byTest[k]) byTest[k] = []; byTest[k].push(r); });
  let rows = '', rowNum = 1;
  Object.entries(byTest).forEach(([tname, tRows]) => {
    rows += `<tr style="background:#1A237E"><td colspan="5" style="padding:7px 12px;font-size:11px;font-weight:800;color:#BBDEFB;letter-spacing:.3px">🔬 ${esc(tname)}</td></tr>`;
    tRows.forEach(r => {
      const paramLabel = r.paramName ? `↳ ${esc(r.paramName)}` : esc(tname);
      rows += `<tr>
        <td style="text-align:center;font-weight:700;color:#94A3B8">${rowNum++}</td>
        <td style="font-size:11px;color:#475569;padding-left:${r.paramName ? '22px' : '12px'}">${paramLabel}</td>
        <td style="font-weight:800;font-family:monospace;font-size:13px">${esc(r.value)}</td>
        <td style="font-size:10px">${esc(r.unit || '—')}</td>
        <td style="font-size:10px">${esc(r.norm || '—')}</td></tr>`;
    });
  });
  document.getElementById('rptBody').innerHTML = rows;
}

// ═══════════════════════════════════════
// SEARCH HELPERS
// ═══════════════════════════════════════
function filterPatList(srchId, selId) {
  const q = gV(srchId).toLowerCase();
  const sel = document.getElementById(selId);
  const pats = ld('patients').filter(p => !CU || CU.role !== 'doctor' || p.ownerId === CU.id);
  const filtered = pats.filter(p => (p.name + p.id).toLowerCase().includes(q));
  
  const currentVal = sel.value;
  sel.innerHTML = '<option value="">' + (q ? `Matching "${q}"...` : 'Select patient...') + '</option>' + 
    filtered.map(p => `<option value="${esc(p.name)}" data-pid="${esc(p.id)}" ${p.name === currentVal ? 'selected' : ''}>${esc(p.name)} (${esc(p.id)})</option>`).join('');
}

function filterOrdList(srchId, selId) {
  const q = gV(srchId).toLowerCase();
  const sel = document.getElementById(selId);
  const ords = ld('orders').filter(o => o.status !== 'Cancelled');
  const filtered = ords.filter(o => (o.id + o.patient).toLowerCase().includes(q));
  
  const currentVal = sel.value;
  sel.innerHTML = '<option value="">' + (q ? `Matching "${q}"...` : '— Select Order —') + '</option>' + 
    filtered.map(o => `<option value="${esc(o.id)}" ${o.id === currentVal ? 'selected' : ''}>${esc(o.id)} — ${esc(o.patient)} (${getTests(o).length} tests)</option>`).join('');
}

// ═══════════════════════════════════════
// EXCEL EXPORT (role-based + auth)
// ═══════════════════════════════════════
function exportExcel() {
  promptAuth('🔐 Confirm Export', 'Enter your credentials to export data.', () => {
    try {
      const wb = XLSX.utils.book_new();
      let _sheetCount = 0; // Bug fix: track sheets added so we can warn if workbook is empty
      if (CU.role === 'admin') {
        // Full system export
        const pats = ld('patients').map(p => ({ 'Patient ID': p.id, 'Full Name': p.name, 'DOB': p.dob, 'Age': p.age, 'Gender': p.gender, 'Blood Type': p.blood, 'Phone': p.phone, 'Email': p.email, 'Address': p.address, 'Insurance': p.ins, 'Type': p.type || 'regular', 'Added By': p.ownerName, 'Reg. Date': p.date, 'Notes': p.notes }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pats), 'Patients');
        const users = ld('users').map(u => ({ 'User ID': u.id, 'Username': u.username, 'Role': u.role, 'Full Name': u.name, 'Specialty': u.specialty, 'Department': u.dept, 'Lab Dept': u.labDept, 'Phone': u.phone, 'Email': u.email, 'Status': u.status, 'Created': u.created }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(users), 'Users');
        const tests = ld('tests').map(t => ({ 'Test ID': t.id, 'Name': t.name, 'Department': t.dept, 'Sample': t.sample, 'Price(₪)': t.price, 'TAT(h)': t.tat, 'Unit': t.unit, 'Normal Range': t.norm, 'Normal Male': t.normM, 'Normal Female': t.normF, 'Method': t.method, 'Status': t.status, 'Params Count': t.params?.length || 0, 'Notes': t.notes }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(tests), 'Tests');
        const ords = ld('orders').map(o => ({ 'Order ID': o.id, 'Date': o.date, 'Patient': o.patient, 'Patient ID': o.patientId, 'Doctor': o.doctor, 'Doctor ID': o.doctorId, 'Tests': getTests(o).join(' | '), 'Priority': o.priority, 'Status': o.status, 'Shift ID': o.shiftId, 'Notes': o.notes }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ords), 'Orders');
        const res = ld('results').map(r => ({ 'Result ID': r.id, 'Date': r.date, 'Order ID': r.orderId, 'Patient': r.patient, 'Test ID': r.testId, 'Test Name': r.testName, 'Parameter': r.paramName, 'Value': r.value, 'Unit': r.unit, 'Normal Range': r.norm, 'Technician': r.tech, 'Comment': r.comment }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(res), 'Results');
        const bill = ld('billing').map(b => ({ 'Invoice ID': b.id, 'Date': b.date, 'Patient': b.patient, 'Doctor': b.doctor, 'Total(₪)': b.total, 'Paid(₪)': b.paid, 'Balance(₪)': Number(b.total || 0) - Number(b.paid || 0), 'Notes': b.notes }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(bill), 'Billing');
        const shifts = ld('shifts').concat(ldOne('curShift') ? [ldOne('curShift')] : []);
        const shData = shifts.map(s => ({ 'Shift ID': s.id, 'Lab User': s.labUser, 'Name': s.labName, 'Start Date': s.startDate, 'Start Time': s.startTime, 'Orders Processed': s.ordersProcessed, 'Results Entered': s.resultsEntered, 'Patients Added': s.patientsAdded }));
        if (shData.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(shData), 'Shifts');
        _sheetCount = wb.SheetNames.length;
        if (!_sheetCount) { toast('Nothing to export — no data found', 'warn'); return; }
        const fname = 'MedLabLIS_AdminExport_' + tod() + '.xlsx'; XLSX.writeFile(wb, fname); toast('✅ Full system export: ' + fname);
      } else if (CU.role === 'doctor') {
        // Doctor: own patients, orders, billing, history
        const myPats = ld('patients').filter(p => p.ownerId === CU.id);
        const patRows = myPats.map(p => ({ 'Patient ID': p.id, 'Full Name': p.name, 'DOB': p.dob, 'Age': p.age, 'Gender': p.gender, 'Blood Type': p.blood, 'Phone': p.phone, 'Email': p.email, 'Address': p.address, 'Insurance': p.ins, 'Reg. Date': p.date, 'Notes': p.notes }));
        if (patRows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(patRows), 'My Patients');
        const myOrds = ld('orders').filter(o => o.doctorId === CU.id).map(o => ({ 'Order ID': o.id, 'Date': o.date, 'Patient': o.patient, 'Tests': getTests(o).join(' | '), 'Priority': o.priority, 'Status': o.status, 'Notes': o.notes }));
        if (myOrds.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(myOrds), 'My Orders');
        const myRes = ld('results').filter(r => ld('orders').find(o => o.id === r.orderId && o.doctorId === CU.id)).map(r => ({ 'Result ID': r.id, 'Date': r.date, 'Order ID': r.orderId, 'Patient': r.patient, 'Test Name': r.testName, 'Parameter': r.paramName, 'Value': r.value, 'Unit': r.unit, 'Normal Range': r.norm }));
        if (myRes.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(myRes), 'My Results');
        const myBill = ld('billing').filter(b => b.doctorId === CU.id).map(b => ({ 'Invoice ID': b.id, 'Date': b.date, 'Patient': b.patient, 'Total(₪)': b.total, 'Paid(₪)': b.paid, 'Balance(₪)': Number(b.total || 0) - Number(b.paid || 0), 'Notes': b.notes }));
        if (myBill.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(myBill), 'My Billing');
        // History tables
        const histCols = getHistCols();
        if (histCols.length) {
          const histRows = [];
          myPats.forEach(p => {
            getHistData(p.id).forEach(row => {
              const obj = { 'Patient ID': p.id, 'Patient Name': p.name, 'Entry Date': row._date || '' };
              histCols.forEach(c => { obj[c.name] = row[c.name] || ''; });
              histRows.push(obj);
            });
          });
          if (histRows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(histRows), 'Patient History');
        }
        _sheetCount = wb.SheetNames.length;
        if (!_sheetCount) { toast('Nothing to export — no data found', 'warn'); return; }
        const fname = 'MedLabLIS_Dr_' + CU.username + '_' + tod() + '.xlsx'; XLSX.writeFile(wb, fname); toast('✅ Doctor export: ' + fname);
      } else if (CU.role === 'lab') {
        // Lab: orders, results, shifts
        const labOrds = ld('orders').map(o => ({ 'Order ID': o.id, 'Date': o.date, 'Patient': o.patient, 'Doctor': o.doctor, 'Tests': getTests(o).join(' | '), 'Priority': o.priority, 'Status': o.status, 'Shift ID': o.shiftId, 'Notes': o.notes }));
        if (labOrds.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(labOrds), 'Orders');
        const labRes = ld('results').map(r => ({ 'Result ID': r.id, 'Date': r.date, 'Order ID': r.orderId, 'Patient': r.patient, 'Test Name': r.testName, 'Parameter': r.paramName, 'Value': r.value, 'Unit': r.unit, 'Normal Range': r.norm, 'Technician': r.tech, 'Comment': r.comment }));
        if (labRes.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(labRes), 'Results');
        const shifts = ld('shifts').concat(ldOne('curShift') ? [ldOne('curShift')] : []);
        const shData = shifts.map(s => ({ 'Shift ID': s.id, 'Lab User': s.labUser, 'Name': s.labName, 'Start Date': s.startDate, 'Start Time': s.startTime, 'Orders Processed': s.ordersProcessed, 'Results Entered': s.resultsEntered, 'Patients Added': s.patientsAdded }));
        if (shData.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(shData), 'Shifts');
        _sheetCount = wb.SheetNames.length;
        if (!_sheetCount) { toast('Nothing to export — no data found', 'warn'); return; }
        const fname = 'MedLabLIS_Lab_' + CU.username + '_' + tod() + '.xlsx'; XLSX.writeFile(wb, fname); toast('✅ Lab export: ' + fname);
      }
    } catch (e) { toast('Export error: ' + e.message, 'error'); }
  });
}

// ═══════════════════════════════════════
// EXCEL IMPORT (role-based + auth)
// ═══════════════════════════════════════
function importExcel() {
  promptAuth('🔐 Confirm Import', 'Enter your credentials to import data.', () => {
    const inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.xlsx,.xls';
    inp.onchange = e => {
      const file = e.target.files[0]; if (!file) return;
      const reader = new FileReader();
      reader.onload = ev => {
        try {
          const wb = XLSX.read(ev.target.result, { type: 'binary' }); let imported = 0;
          if (CU.role === 'admin') {
            if (wb.SheetNames.includes('Patients')) { const rows = XLSX.utils.sheet_to_json(wb.Sheets['Patients']); const valid = rows.filter(r => r['Patient ID'] && r['Full Name']); if (valid.length) { const mapped = valid.map(r => ({ id: String(r['Patient ID'] || ''), name: String(r['Full Name'] || ''), dob: String(r['DOB'] || ''), age: r['Age'] || '', gender: String(r['Gender'] || ''), blood: String(r['Blood Type'] || ''), phone: String(r['Phone'] || ''), email: String(r['Email'] || ''), address: String(r['Address'] || ''), ins: String(r['Insurance'] || ''), type: String(r['Type'] || 'regular'), ownerName: String(r['Added By'] || ''), date: String(r['Reg. Date'] || tod()), notes: String(r['Notes'] || ''), ownerId: '', ownerRole: '' })); const ex = ld('patients'); mapped.forEach(row => { const i = ex.findIndex(x => x.id === row.id); if (i >= 0) ex[i] = row; else ex.push(row); }); sv('patients', ex); imported += mapped.length; } }
            if (wb.SheetNames.includes('Orders')) { const rows = XLSX.utils.sheet_to_json(wb.Sheets['Orders']); const valid = rows.filter(r => r['Order ID'] && r['Patient']); if (valid.length) { const mapped = valid.map(r => ({ id: String(r['Order ID'] || ''), date: String(r['Date'] || tod()), patient: String(r['Patient'] || ''), patientId: String(r['Patient ID'] || ''), doctor: String(r['Doctor'] || ''), doctorId: String(r['Doctor ID'] || ''), tests: (String(r['Tests'] || '')).split(' | ').filter(Boolean), priority: String(r['Priority'] || 'Routine'), status: String(r['Status'] || 'Pending'), shiftId: String(r['Shift ID'] || ''), notes: String(r['Notes'] || '') })); const ex = ld('orders'); mapped.forEach(row => { const i = ex.findIndex(x => x.id === row.id); if (i >= 0) ex[i] = row; else ex.push(row); }); sv('orders', ex); imported += mapped.length; } }
            if (wb.SheetNames.includes('Results')) { const rows = XLSX.utils.sheet_to_json(wb.Sheets['Results']); const valid = rows.filter(r => r['Result ID'] && r['Test Name']); if (valid.length) { const mapped = valid.map(r => ({ id: String(r['Result ID'] || ''), date: String(r['Date'] || tod()), orderId: String(r['Order ID'] || ''), patient: String(r['Patient'] || ''), testId: String(r['Test ID'] || ''), testName: String(r['Test Name'] || ''), paramName: String(r['Parameter'] || ''), value: String(r['Value'] || ''), unit: String(r['Unit'] || ''), norm: String(r['Normal Range'] || ''), tech: String(r['Technician'] || ''), comment: String(r['Comment'] || '') })); const ex = ld('results'); mapped.forEach(row => { const i = ex.findIndex(x => x.id === row.id); if (i >= 0) ex[i] = row; else ex.push(row); }); sv('results', ex); imported += mapped.length; } }
            if (wb.SheetNames.includes('Billing')) { const rows = XLSX.utils.sheet_to_json(wb.Sheets['Billing']); const valid = rows.filter(r => r['Invoice ID'] && r['Patient']); if (valid.length) { const mapped = valid.map(r => ({ id: String(r['Invoice ID'] || ''), date: String(r['Date'] || tod()), patient: String(r['Patient'] || ''), doctor: String(r['Doctor'] || ''), total: Number(r['Total(₪)']) || 0, paid: Number(r['Paid(₪)']) || 0, notes: String(r['Notes'] || '') })); const ex = ld('billing'); mapped.forEach(row => { const i = ex.findIndex(x => x.id === row.id); if (i >= 0) ex[i] = row; else ex.push(row); }); sv('billing', ex); imported += mapped.length; } }
            toast(`✅ Admin import complete — ${imported} records merged`); renderAdmDash();
          } else if (CU.role === 'doctor') {
            if (wb.SheetNames.includes('My Patients')) { const rows = XLSX.utils.sheet_to_json(wb.Sheets['My Patients']); const valid = rows.filter(r => r['Patient ID'] && r['Full Name']); if (valid.length) { const mapped = valid.map(r => ({ id: String(r['Patient ID'] || ''), name: String(r['Full Name'] || ''), dob: String(r['DOB'] || ''), age: r['Age'] || '', gender: String(r['Gender'] || ''), blood: String(r['Blood Type'] || ''), phone: String(r['Phone'] || ''), email: String(r['Email'] || ''), address: String(r['Address'] || ''), ins: String(r['Insurance'] || ''), type: 'regular', ownerName: CU.name, date: String(r['Reg. Date'] || tod()), notes: String(r['Notes'] || ''), ownerId: CU.id, ownerRole: 'doctor' })); const ex = ld('patients'); mapped.forEach(row => { const i = ex.findIndex(x => x.id === row.id); if (i >= 0) ex[i] = row; else ex.push(row); }); sv('patients', ex); imported += mapped.length; } }
            if (wb.SheetNames.includes('My Orders')) { const rows = XLSX.utils.sheet_to_json(wb.Sheets['My Orders']); const valid = rows.filter(r => r['Order ID'] && r['Patient']); if (valid.length) { const mapped = valid.map(r => ({ id: String(r['Order ID'] || ''), date: String(r['Date'] || tod()), patient: String(r['Patient'] || ''), patientId: '', doctorId: CU.id, doctor: CU.name, tests: (String(r['Tests'] || '')).split(' | ').filter(Boolean), priority: String(r['Priority'] || 'Routine'), status: String(r['Status'] || 'Pending'), shiftId: '', notes: String(r['Notes'] || '') })); const ex = ld('orders'); mapped.forEach(row => { const i = ex.findIndex(x => x.id === row.id); if (i >= 0) ex[i] = row; else ex.push(row); }); sv('orders', ex); imported += mapped.length; } }
            toast(`✅ Doctor import complete — ${imported} records merged`); renderDocDash();
          } else if (CU.role === 'lab') {
            if (wb.SheetNames.includes('Results')) { const rows = XLSX.utils.sheet_to_json(wb.Sheets['Results']); const valid = rows.filter(r => r['Result ID'] && r['Test Name']); if (valid.length) { const mapped = valid.map(r => ({ id: String(r['Result ID'] || ''), date: String(r['Date'] || tod()), orderId: String(r['Order ID'] || ''), patient: String(r['Patient'] || ''), testId: String(r['Test ID'] || ''), testName: String(r['Test Name'] || ''), paramName: String(r['Parameter'] || ''), value: String(r['Value'] || ''), unit: String(r['Unit'] || ''), norm: String(r['Normal Range'] || ''), tech: String(r['Technician'] || CU.username), comment: String(r['Comment'] || '') })); const ex = ld('results'); mapped.forEach(row => { const i = ex.findIndex(x => x.id === row.id); if (i >= 0) ex[i] = row; else ex.push(row); }); sv('results', ex); imported += mapped.length; } }
            if (wb.SheetNames.includes('Orders')) { const rows = XLSX.utils.sheet_to_json(wb.Sheets['Orders']); const valid = rows.filter(r => r['Order ID'] && r['Patient']); if (valid.length) { const mapped = valid.map(r => ({ id: String(r['Order ID'] || ''), date: String(r['Date'] || tod()), patient: String(r['Patient'] || ''), patientId: '', doctor: String(r['Doctor'] || ''), doctorId: '', tests: (String(r['Tests'] || '')).split(' | ').filter(Boolean), priority: String(r['Priority'] || 'Routine'), status: String(r['Status'] || 'Pending'), shiftId: String(r['Shift ID'] || ''), notes: String(r['Notes'] || '') })); const ex = ld('orders'); mapped.forEach(row => { const i = ex.findIndex(x => x.id === row.id); if (i >= 0) ex[i] = row; else ex.push(row); }); sv('orders', ex); imported += mapped.length; } }
            toast(`✅ Lab import complete — ${imported} records merged`); renderLabDash();
          }
        } catch (err) { toast('Import failed: ' + err.message, 'error'); }
      };
      reader.readAsBinaryString(file);
    };
    inp.click();
  });
}

// ═══════════════════════════════════════
// STARTUP
// ═══════════════════════════════════════
initData().then(() => { checkSession(); });
setRole('doctor');

