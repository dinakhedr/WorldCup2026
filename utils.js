// ============================================================
// UTILS.JS — World Cup 2026 Tracker
// Shared constants, auth, GAPI, helpers, score modal
// ============================================================

// ── Constants ─────────────────────────────────────────────
const CLIENT_ID      = '932472749275-3s12tjsttbppmc96gf5rdl6orrpj0o5c.apps.googleusercontent.com';
const SPREADSHEET_ID = '1a-VzpWXZMrQZqwBLV_BstNREArnC8mgxuvSdz2MyjRM';
const SCOPES         = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file';
const DISCOVERY_DOCS = [
  'https://sheets.googleapis.com/$discovery/rest?version=v4',
  'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'
];

const SHEET_GM       = 'Group Matches';
const SHEET_MATCHES  = 'Matches';
const SHEET_STANDING = 'Standings';
const SHEET_KO       = 'Knockout Results';

// ── Group colors ───────────────────────────────────────────
const GRP_COLORS = {
  A:{ bg:'#d1fae5', txt:'#065f46', acc:'#1B6E3F' },
  B:{ bg:'#dbeafe', txt:'#1e40af', acc:'#1a56a0' },
  C:{ bg:'#fef3c7', txt:'#92400e', acc:'#b45309' },
  D:{ bg:'#fee2e2', txt:'#991b1b', acc:'#dc2626' },
  E:{ bg:'#ede9fe', txt:'#5b21b6', acc:'#7c3aed' },
  F:{ bg:'#cffafe', txt:'#155e75', acc:'#0e7490' },
  G:{ bg:'#ffedd5', txt:'#9a3412', acc:'#ea580c' },
  H:{ bg:'#fce7f3', txt:'#9d174d', acc:'#db2777' },
  I:{ bg:'#e0e7ff', txt:'#3730a3', acc:'#4338ca' },
  J:{ bg:'#e0f2fe', txt:'#0c4a6e', acc:'#0284c7' },
  K:{ bg:'#dcfce7', txt:'#166534', acc:'#16a34a' },
  L:{ bg:'#fef9c3', txt:'#854d0e', acc:'#ca8a04' },
};

const KO_STAGES = ['Round of 32','Round of 16','Quarter-Final','Semi-Final','Bronze Medal','Final'];

const GROUPS_TEAMS = {
  A:['Mexico 🇲🇽','South Africa 🇿🇦','South Korea 🇰🇷','Czechia 🇨🇿'],
  B:['Canada 🇨🇦','Bosnia & Herz. 🇧🇦','Qatar 🇶🇦','Switzerland 🇨🇭'],
  C:['Brazil 🇧🇷','Morocco 🇲🇦','Haiti 🇭🇹','Scotland 🏴󠁧󠁢󠁳󠁣󠁴󠁿'],
  D:['USA 🇺🇸','Paraguay 🇵🇾','Australia 🇦🇺','Turkiye 🇹🇷'],
  E:['Germany 🇩🇪','Curacao 🇨🇼','Ivory Coast 🇨🇮','Ecuador 🇪🇨'],
  F:['Netherlands 🇳🇱','Japan 🇯🇵','Sweden 🇸🇪','Tunisia 🇹🇳'],
  G:['Belgium 🇧🇪','Egypt 🇪🇬','Iran 🇮🇷','New Zealand 🇳🇿'],
  H:['Spain 🇪🇸','Cape Verde 🇨🇻','Saudi Arabia 🇸🇦','Uruguay 🇺🇾'],
  I:['France 🇫🇷','Senegal 🇸🇳','Iraq 🇮🇶','Norway 🇳🇴'],
  J:['Argentina 🇦🇷','Algeria 🇩🇿','Austria 🇦🇹','Jordan 🇯🇴'],
  K:['Portugal 🇵🇹','DR Congo 🇨🇩','Uzbekistan 🇺🇿','Colombia 🇨🇴'],
  L:['England 🏴󠁧󠁢󠁥󠁮󠁧󠁿','Croatia 🇭🇷','Ghana 🇬🇭','Panama 🇵🇦'],
};

const NAV_TABS = [
  { id:'home',      label:'Home',     icon:'🏠', href:'Home.html'      },
  { id:'today',     label:'Today',    icon:'📅', href:'Today.html'     },
  { id:'matches',   label:'Matches',  icon:'📋', href:'Matches.html'   },
  { id:'groups',    label:'Groups',   icon:'👥', href:'Groups.html'    },
  { id:'knockout',  label:'Knockout', icon:'⚡', href:'Knockout.html'  },
  { id:'standings', label:'Standings',icon:'📊', href:'Standings.html' },
];

// ── Token helpers ──────────────────────────────────────────
function saveAccessToken(t)  { sessionStorage.setItem('wc_access_token', t); }
function getAccessToken()    { return sessionStorage.getItem('wc_access_token'); }
function clearAccessToken()  { sessionStorage.removeItem('wc_access_token'); }

function getEmailFromJWT(jwt) {
  try {
    const json = decodeURIComponent(
      atob(jwt.split('.')[1].replace(/-/g,'+').replace(/_/g,'/'))
        .split('').map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2)).join('')
    );
    return JSON.parse(json).email || null;
  } catch(e) { return null; }
}

function isIOSSafari() {
  return (
    /iP(hone|od|ad)/.test(navigator.userAgent) &&
    /WebKit/.test(navigator.userAgent) &&
    !/CriOS|FxiOS/.test(navigator.userAgent)
  ) || window.navigator.standalone === true;
}

// ── GAPI init ──────────────────────────────────────────────
async function initGapiClient() {
  await new Promise(resolve => {
    if (gapi.client?.sheets) { resolve(); return; }
    gapi.load('client', resolve);
  });
  await gapi.client.init({ apiKey: '', discoveryDocs: DISCOVERY_DOCS });
}

// ── Auth flow ──────────────────────────────────────────────
// Each page calls initPageAuth({ onReady, activePage })
// onReady() is called once a valid token is confirmed
// activePage is the nav tab id to highlight

async function initPageAuth({ onReady, activePage = '' }) {
  const jwt = localStorage.getItem('wc_google_token');
  if (!jwt) { window.location.href = 'Home.html'; return; }

  renderBottomNav(activePage);
  await initGapiClient();

  const saved = getAccessToken();
  if (saved) {
    gapi.client.setToken({ access_token: saved });
    try {
      // Validate token with a lightweight call
      await gapi.client.sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID, range: `${SHEET_GM}!A1`
      });
      await onReady();
      return;
    } catch(e) { clearAccessToken(); }
  }

  // Need a new token
  const doReq = () => {
    google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID, scope: SCOPES, prompt: '',
      callback: async (resp) => {
        if (!resp.access_token) { window.location.href = 'Home.html'; return; }
        saveAccessToken(resp.access_token);
        gapi.client.setToken({ access_token: resp.access_token });
        await onReady();
      }
    }).requestAccessToken();
  };

  if (isIOSSafari()) { showIOSTapScreen(doReq); }
  else { doReq(); }
}

// iOS tap-to-connect overlay
function showIOSTapScreen(onTap) {
  const box = document.getElementById('loadingBox');
  if (!box) return;
  box.innerHTML = `
    <span class="screen-trophy">🏆</span>
    <div class="screen-title">Welcome Back</div>
    <div class="screen-sub">Tap to connect to Google</div>
    <button class="tap-btn" id="iosTapBtn">Connect to Google</button>`;
  document.getElementById('iosTapBtn').addEventListener('click', () => {
    box.innerHTML = `<span class="screen-trophy">🏆</span><div class="screen-title">Connecting…</div><div class="spinner"></div>`;
    onTap();
  });
}

// ── Bottom nav renderer ────────────────────────────────────
function renderBottomNav(activePage, mountId = 'bottomNavMount') {
  const mount = document.getElementById(mountId);
  if (!mount) return;
  const tabs = NAV_TABS.map(t => `
    <a class="nav-tab${t.id === activePage ? ' active' : ''}" href="${t.href}">
      <span class="nav-tab-icon">${t.icon}</span>${t.label}
    </a>`).join('');
  mount.innerHTML = `<nav class="bottom-nav"><div class="nav-tabs">${tabs}</div></nav>`;
}

// ── Screen helpers ─────────────────────────────────────────
function showLoadingScreen(msg = 'Loading…') {
  const s = document.getElementById('loadingScreen');
  const m = document.getElementById('mainApp');
  if (s) s.classList.add('active');
  if (m) m.style.display = 'none';
  const lbl = document.getElementById('loadingMsg');
  if (lbl) lbl.textContent = msg;
}

function showMainApp() {
  const s = document.getElementById('loadingScreen');
  const m = document.getElementById('mainApp');
  if (s) s.classList.remove('active');
  if (m) m.style.display = 'block';
}

// ── Toast ──────────────────────────────────────────────────
let _toastTimer = null;
function showToast(msg, type = 'success') {
  const t = document.getElementById('toast');
  if (!t) return;
  if (_toastTimer) { clearTimeout(_toastTimer); _toastTimer = null; }
  t.className = `toast ${type}`;
  t.textContent = msg;
  void t.offsetHeight;
  t.classList.add('show');
  _toastTimer = setTimeout(() => t.classList.remove('show'), 2600);
}

// ── Date helpers ───────────────────────────────────────────
// ── 12-hour time formatter ─────────────────────────────────
// Input: "22:00" → "10:00 PM"  |  "05:00" → "5:00 AM"
function formatTime12(time24) {
  if (!time24 || !time24.includes(':')) return time24 || '';
  const [hStr, mStr] = time24.split(':');
  let h = parseInt(hStr, 10);
  const m = mStr || '00';
  const ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12;
  if (h === 0) h = 12;
  return `${h}:${m} ${ampm}`;
}

function getTodayCairo() {
  const now   = new Date();
  const cairo = new Date(now.toLocaleString('en-US', { timeZone: 'Africa/Cairo' }));
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const d = String(cairo.getDate()).padStart(2,'0');
  const m = months[cairo.getMonth()];
  const y = String(cairo.getFullYear()).slice(-2);
  return `${d}-${m}-${y}`;
}

function getTomorrowCairo() {
  const now   = new Date();
  const cairo = new Date(now.toLocaleString('en-US', { timeZone: 'Africa/Cairo' }));
  cairo.setDate(cairo.getDate() + 1);
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const d = String(cairo.getDate()).padStart(2,'0');
  const m = months[cairo.getMonth()];
  const y = String(cairo.getFullYear()).slice(-2);
  return `${d}-${m}-${y}`;
}

function formatDateLabel(d) {
  try {
    const parts = d.split('-');
    if (parts.length !== 3) return d;
    const dt = new Date(`${parts[1]} ${parts[0]}, 20${parts[2]}`);
    if (isNaN(dt)) return d;
    const days   = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `${days[dt.getDay()]}, ${parseInt(parts[0])} ${months[dt.getMonth()]} 2026`;
  } catch(e) { return d; }
}

// ── Row parser (Group Matches sheet) ──────────────────────
// Cols: A=#  B=Grp  C=Date  D=Time  E=Home  F=Hg  G=Ag  H=PenH  I=PenA  J=Away  K=Venue  L=City  M=Country  N=Winner  O=Status
function parseGMRow(row, rowIndex) {
  const num = parseInt(row[0]);
  if (isNaN(num) || num < 1) return null;
  const grp  = (row[1] || '').trim();
  const isKO = KO_STAGES.includes(grp);
  return {
    rowIndex,
    num,
    group:     grp,
    isKO,
    date:      (row[2]  || '').trim(),
    time:      (row[3]  || '').trim(),
    home:      (row[4]  || '').trim(),
    homeScore: row[5] !== undefined && row[5] !== '' ? String(row[5]) : '',
    awayScore: row[6] !== undefined && row[6] !== '' ? String(row[6]) : '',
    penH:      row[7] !== undefined && row[7] !== '' ? String(row[7]) : '',
    penA:      row[8] !== undefined && row[8] !== '' ? String(row[8]) : '',
    away:      (row[9]  || '').trim(),
    venue:     (row[10] || '').trim(),
    city:      (row[11] || '').trim(),
    country:   (row[12] || '').trim(),
    winner:    (row[13] || '-').trim(),
    status:    (row[14] || 'Upcoming').trim(),
    homeYellow: parseInt(r[15]) || 0,
    homeRed:    parseInt(r[16]) || 0,
    awayYellow: parseInt(r[17]) || 0,
    awayRed:    parseInt(r[18]) || 0,
  };
}

// ── Load all Group Matches rows (with 60s cache) ───────────
const _CACHE_KEY = 'wc_matches_cache';
const _CACHE_TTL = 60000; // 60 seconds

function _invalidateMatchCache() {
  sessionStorage.removeItem(_CACHE_KEY);
}

async function loadAllMatches(forceRefresh = false) {
  if (!forceRefresh) {
    try {
      const cached = sessionStorage.getItem(_CACHE_KEY);
      if (cached) {
        const { ts, data } = JSON.parse(cached);
        if (Date.now() - ts < _CACHE_TTL) return data;
      }
    } catch(e) { /* ignore bad cache */ }
  }

  let res;
  try {
    res = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_GM}!A:O`
    });
  } catch(e) {
    // Token likely expired — clear it so next call re-auths
    if (e.status === 401 || (e.result && e.result.error && e.result.error.code === 401)) {
      clearAccessToken();
    }
    throw e;
  }

  const rows = res.result.values || [];
  const matches = [];
  rows.forEach((row, idx) => {
    const m = parseGMRow(row, idx);
    if (m) matches.push(m);
  });

  try {
    sessionStorage.setItem(_CACHE_KEY, JSON.stringify({ ts: Date.now(), data: matches }));
  } catch(e) { /* quota exceeded — ignore */ }

  return matches;
}

// ── Match card HTML ────────────────────────────────────────
// Used across Matches, Today, Groups, Knockout pages
function matchCardHTML(m, opts = {}) {
  const { showDate = false, onclick = `openScoreModal(${m.num})` } = opts;
  const played  = m.status === 'Played';
  const isKO    = m.isKO;
  const grpCol  = GRP_COLORS[m.group] || { bg:'#f3f4f6', txt:'#374151', acc:'#6b7280' };
  const acc     = isKO ? '#ea580c' : grpCol.acc;

  let badgeBg, badgeTxt, badgeLabel;
  if (isKO) {
    badgeBg = '#fff3cd'; badgeTxt = '#7B5800';
    badgeLabel = m.group === 'Final' ? '🏆 Final'
               : m.group === 'Bronze Medal' ? '🥉 Bronze'
               : `⚡ ${m.group}`;
  } else {
    badgeBg = grpCol.bg; badgeTxt = grpCol.txt;
    badgeLabel = `Group ${m.group}`;
  }

  const timeStr = formatTime12(m.time);

  let scoreHTML;
  if (played) {
    scoreHTML = `<div class="score-box played">
      <span class="score-digit played">${m.homeScore}</span>
      <span class="score-sep">:</span>
      <span class="score-digit played">${m.awayScore}</span>
    </div>`;
  } else {
    scoreHTML = `<div class="score-box"><span class="score-vs">${timeStr}</span></div>`;
  }

  const homeWin  = played && m.winner === m.home;
  const awayWin  = played && m.winner === m.away;
  const homeClass = `team-name${homeWin ? ' winner-team' : ''}`;
  const awayClass = `team-name away${awayWin ? ' winner-team' : ''}`;
  const venueStr  = [m.venue, m.city, m.country].filter(Boolean).join(' · ');
  const penBadge  = (played && m.penH !== '' && m.penA !== '')
    ? `<span class="pen-badge">Pen ${m.penH}–${m.penA}</span>` : '';
  const dateLine  = showDate ? `<span class="match-date-inline">${m.date} · </span>` : '';

  return `
    <div class="match-card${played ? ' played' : ''}" style="--card-acc:${acc}" onclick="${onclick}">
      <div class="card-inner">
        <div class="card-top">
          <span class="stage-badge" style="background:${badgeBg};color:${badgeTxt}">${badgeLabel}</span>
          <span class="match-time">${played
            ? '<span class="played-dot"></span>FT'
            : `${dateLine}${timeStr} Cairo`}</span>
        </div>
        <div class="teams-row">
          <div class="${homeClass}">${escHtml(m.home)}</div>
          ${scoreHTML}
          <div class="${awayClass}">${escHtml(m.away)}</div>
        </div>
      </div>
      <div class="card-bottom">
        <span class="venue-txt">📍 ${venueStr}</span>
        ${penBadge}
      </div>
    </div>`;
}

// ── Score modal ────────────────────────────────────────────
// Pages include the modal HTML via renderScoreModalHTML()
// and call openScoreModal(matchNum) / closeScoreModal()
// The page must expose window._allMatches and window._onScoreSaved()

function renderScoreModalHTML() {
  return `
  <div class="modal-backdrop" id="scoreModal" onclick="if(event.target===this)closeScoreModal()">
    <div class="modal">
      <div class="modal-handle"></div>
      <div class="modal-match-info">
        <div class="modal-badge" id="modalBadge"></div>
        <div class="modal-teams" id="modalTeams"></div>
        <div class="modal-meta"  id="modalMeta"></div>
      </div>
      <div class="score-inputs">
        <input class="score-inp" id="inpHome" type="number" min="0" max="30" placeholder="0" inputmode="numeric">
        <div class="score-vs-lbl">vs</div>
        <input class="score-inp" id="inpAway" type="number" min="0" max="30" placeholder="0" inputmode="numeric">
      </div>
      <div class="team-lbls">
        <div class="team-lbl" id="lblHome"></div>
        <div></div>
        <div class="team-lbl" id="lblAway"></div>
      </div>
      <div class="pen-section" id="penSection" style="display:none;">
        <div class="pen-title">🥅 Penalties (if scores level)</div>
        <div class="pen-inputs">
          <input class="pen-inp" id="inpPenH" type="number" min="0" max="20" placeholder="—" inputmode="numeric">
          <div class="score-vs-lbl" style="color:#92400e;">vs</div>
          <input class="pen-inp" id="inpPenA" type="number" min="0" max="20" placeholder="—" inputmode="numeric">
        </div>
      </div>
      <div class="modal-footer">
        <button class="btn-cancel" onclick="closeScoreModal()">Cancel</button>
        <button class="btn-clear"  id="btnClear" onclick="clearScoreResult()">Clear</button>
        <button class="btn-save"   id="btnSave" onclick="saveScoreResult()">Save Result</button>
      </div>
    </div>
  </div>`;
}

let _editingMatch = null;

function openScoreModal(matchNum) {
  const matches = window._allMatches || [];
  const m = matches.find(x => x.num === matchNum);
  if (!m) return;
  _editingMatch = m;

  const isKO    = m.isKO;
  const grpCol  = GRP_COLORS[m.group] || { bg:'#f3f4f6', txt:'#374151' };
  const badgeBg = isKO ? '#fff3cd' : grpCol.bg;
  const badgeTxt= isKO ? '#7B5800' : grpCol.txt;
  const badgeLbl= isKO ? `⚡ ${m.group}` : `Group ${m.group}`;

  const badge = document.getElementById('modalBadge');
  badge.textContent        = badgeLbl;
  badge.style.background   = badgeBg;
  badge.style.color        = badgeTxt;

  document.getElementById('modalTeams').textContent = `${m.home} vs ${m.away}`;
  document.getElementById('modalMeta').textContent  = `${m.date} · ${formatTime12(m.time)} Cairo · ${m.venue}, ${m.city}`;
  document.getElementById('lblHome').textContent    = m.home;
  document.getElementById('lblAway').textContent    = m.away;

  document.getElementById('inpHome').value  = m.homeScore !== '' ? m.homeScore : '';
  document.getElementById('inpAway').value  = m.awayScore !== '' ? m.awayScore : '';
  document.getElementById('inpPenH').value  = m.penH !== '' ? m.penH : '';
  document.getElementById('inpPenA').value  = m.penA !== '' ? m.penA : '';

  document.getElementById('penSection').style.display = isKO ? 'block' : 'none';
  document.getElementById('scoreModal').classList.add('open');
  setTimeout(() => document.getElementById('inpHome').focus(), 320);
}

function closeScoreModal() {
  document.getElementById('scoreModal').classList.remove('open');
  _editingMatch = null;
  // Reset clear button confirming state
  const btn = document.getElementById('btnClear');
  if (btn) {
    btn.dataset.confirming = 'false';
    btn.textContent = 'Clear';
    btn.style.background = '';
    btn.style.color = '';
  }
}

async function saveScoreResult() {
  if (!_editingMatch) return;
  const hv = document.getElementById('inpHome').value;
  const av = document.getElementById('inpAway').value;
  if (hv === '' || av === '') return showToast('Enter both scores', 'error');
  const hs = parseInt(hv), as2 = parseInt(av);
  if (isNaN(hs) || isNaN(as2) || hs < 0 || as2 < 0) return showToast('Invalid score', 'error');

  const ph = document.getElementById('inpPenH').value;
  const pa = document.getElementById('inpPenA').value;
  const btn = document.getElementById('btnSave');
  btn.textContent = 'Saving…'; btn.disabled = true;

  try {
    const sheetRow = _editingMatch.rowIndex + 1;
    await gapi.client.sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_GM}!F${sheetRow}:G${sheetRow}`,
      valueInputOption: 'RAW',
      resource: { values: [[hs, as2]] }
    });

    if (_editingMatch.isKO) {
      if (ph !== '' && pa !== '') {
        await gapi.client.sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `${SHEET_GM}!H${sheetRow}:I${sheetRow}`,
          valueInputOption: 'RAW',
          resource: { values: [[parseInt(ph)||0, parseInt(pa)||0]] }
        });
      } else {
        await gapi.client.sheets.spreadsheets.values.clear({
          spreadsheetId: SPREADSHEET_ID,
          range: `${SHEET_GM}!H${sheetRow}:I${sheetRow}`
        });
      }
    }

    closeScoreModal();
    showToast('✅ Result saved!');
    _invalidateMatchCache();
    if (typeof window._onScoreSaved === 'function') await window._onScoreSaved();
  } catch(e) {
    console.error(e);
    showToast('Failed to save', 'error');
  } finally {
    btn.textContent = 'Save Result'; btn.disabled = false;
  }
}

async function clearScoreResult() {
  if (!_editingMatch) return;
  // iOS-safe in-UI confirm — replace Clear button with two inline buttons
  const btn = document.getElementById('btnClear');
  if (btn.dataset.confirming === 'true') {
    // Second tap — actually clear
    btn.textContent = 'Clearing…'; btn.disabled = true;
    const sheetRow = _editingMatch.rowIndex + 1;
    try {
      await gapi.client.sheets.spreadsheets.values.clear({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_GM}!F${sheetRow}:I${sheetRow}`
      });
      _invalidateMatchCache();
      closeScoreModal();
      showToast('Result cleared');
      if (typeof window._onScoreSaved === 'function') await window._onScoreSaved();
    } catch(e) {
      showToast('Failed to clear', 'error');
      btn.textContent = 'Clear'; btn.disabled = false;
      btn.dataset.confirming = 'false';
    }
  } else {
    // First tap — ask for confirmation inline
    btn.dataset.confirming = 'true';
    btn.textContent = 'Tap again to confirm';
    btn.style.background = '#fee2e2';
    btn.style.color = '#991b1b';
    // Auto-reset after 3s if not confirmed
    setTimeout(() => {
      if (btn.dataset.confirming === 'true') {
        btn.dataset.confirming = 'false';
        btn.textContent = 'Clear';
        btn.style.background = '';
        btn.style.color = '';
      }
    }, 3000);
  }
}

// ── Sign out ───────────────────────────────────────────────
function doSignOut() {
  if (!confirm('Sign out of World Cup 2026?')) return;
  clearAccessToken();
  localStorage.removeItem('wc_google_token');
  window.location.href = 'Home.html';
}

// ── Escape HTML ────────────────────────────────────────────
function escHtml(s) {
  return (s||'').replace(/[&<>"']/g, m =>
    ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])
  );
}
