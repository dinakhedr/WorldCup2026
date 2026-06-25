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

const SHEET_GM       = 'Matches';
const SHEET_STANDING = 'Standings';

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
    homeYellow: parseInt(row[15]) || 0,
    homeRed:    parseInt(row[16]) || 0,
    awayYellow: parseInt(row[17]) || 0,
    awayRed:    parseInt(row[18]) || 0,
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
      range: `${SHEET_GM}!A:S`
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
  const isUpcoming = m.status !== 'Played';
  const notifActive = isUpcoming && isNotifScheduled(m.num);
  const bellBtn = isUpcoming
    ? `<button class="bell-btn" data-notif-btn="${m.num}" onclick="event.stopPropagation();toggleMatchNotification(${m.num})" style="opacity:${notifActive ? '1' : '0.5'};">${notifActive ? '🔔' : '🔕'}</button>`
    : '';
  return `
    <div class="match-card${played ? ' played' : ''}" style="--card-acc:${acc}" onclick="${onclick}">
      <div class="card-inner">
        <div class="card-top">
          <span class="stage-badge" style="background:${badgeBg};color:${badgeTxt}">${badgeLabel}</span>
          <span class="match-time" style="display:flex;align-items:center;gap:6px;">${played
            ? '<span class="played-dot"></span>FT'
            : `${dateLine}${timeStr} Cairo`}${bellBtn}</span>
        </div>
        <div class="teams-row">
          <div style="display:flex;flex-direction:column;gap:3px;">
            <div class="${homeClass}">${escHtml(m.home)}</div>
            ${m.status === 'Played' && (m.homeYellow || m.homeRed) ? `<div class="match-cards-home">${m.homeYellow ? `<span class="card-pill">🟨×${m.homeYellow}</span>` : ''}${m.homeRed ? `<span class="card-pill">🟥×${m.homeRed}</span>` : ''}</div>` : ''}
          </div>
          ${scoreHTML}
          <div style="display:flex;flex-direction:column;gap:3px;align-items:flex-end;">
            <div class="${awayClass}">${escHtml(m.away)}</div>
            ${m.status === 'Played' && (m.awayYellow || m.awayRed) ? `<div class="match-cards-away">${m.awayYellow ? `<span class="card-pill">🟨×${m.awayYellow}</span>` : ''}${m.awayRed ? `<span class="card-pill">🟥×${m.awayRed}</span>` : ''}</div>` : ''}
          </div>
        </div>
      </div>
      <div class="card-bottom">
        <span class="venue-txt">📍 ${venueStr}</span>
        ${penBadge}
      </div>
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
      <div class="cards-section">
        <div class="cards-title">🟨 🟥 Cards</div>
        <div class="cards-inputs-row">
          <input class="card-inp yellow" id="inpHomeYellow" type="number" min="0" max="20" placeholder="0" inputmode="numeric">
          <span class="card-emoji">🟨</span>
          <input class="card-inp yellow" id="inpAwayYellow" type="number" min="0" max="20" placeholder="0" inputmode="numeric">
        </div>
        <div class="cards-inputs-row" style="margin-top:8px;">
          <input class="card-inp red" id="inpHomeRed" type="number" min="0" max="20" placeholder="0" inputmode="numeric">
          <span class="card-emoji">🟥</span>
          <input class="card-inp red" id="inpAwayRed" type="number" min="0" max="20" placeholder="0" inputmode="numeric">
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

  document.getElementById('inpHome').value  = m.homeScore !== '' ? m.homeScore : '0';
  document.getElementById('inpAway').value  = m.awayScore !== '' ? m.awayScore : '0';
  document.getElementById('inpPenH').value  = m.penH !== '' ? m.penH : '0';
  document.getElementById('inpPenA').value  = m.penA !== '' ? m.penA : '0';
  document.getElementById('inpHomeYellow').value = m.homeYellow || '0';
  document.getElementById('inpHomeRed').value    = m.homeRed    || '0';
  document.getElementById('inpAwayYellow').value = m.awayYellow || '0';
  document.getElementById('inpAwayRed').value    = m.awayRed    || '0';

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

    const hy = parseInt(document.getElementById('inpHomeYellow').value) || 0;
    const hr = parseInt(document.getElementById('inpHomeRed').value)    || 0;
    const ay = parseInt(document.getElementById('inpAwayYellow').value) || 0;
    const ar = parseInt(document.getElementById('inpAwayRed').value)    || 0;
    await gapi.client.sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_GM}!P${sheetRow}:S${sheetRow}`,
      valueInputOption: 'RAW',
      resource: { values: [[hy, hr, ay, ar]] }
    });
    
    closeScoreModal();
    showToast('✅ Result saved!');
    _invalidateMatchCache();
    await writeResolvedKOTeams();
    if (typeof window._onScoreSaved === 'function') await window._onScoreSaved();
  } catch(e) {
    console.error(e);
    showToast('Failed to save', 'error');
  } finally {
    btn.textContent = 'Save Result'; btn.disabled = false;
  }
}

/* ── Write resolved KO team names back to the sheet ──────────
   Called after every score save. Resolves R32 slots from current
   standings and writes team names (with flags) to cols E & J.
   Only writes cells that have changed from their current value.  */
async function writeResolvedKOTeams() {
  try {
    const allMatches = await loadAllMatches(true);
    const { standings, thirdPlaceTeams } = buildAllStandings(allMatches);

    // Build list of sheet updates needed
    const updates = [];

    for (const [num, seeding] of Object.entries(R32_SEEDING)) {
      const matchNum = parseInt(num);
      const m = allMatches.find(x => x.num === matchNum);
      if (!m) continue;

      const resolvedHome = resolveR32Slot(seeding.home, standings, thirdPlaceTeams, allMatches);
      const resolvedAway = resolveR32Slot(seeding.away, standings, thirdPlaceTeams, allMatches);

      const sheetRow = m.rowIndex + 1;

      // Only write if resolved and different from current cell value
      if (resolvedHome && resolvedHome !== m.home) {
        updates.push({
          range: `${SHEET_GM}!E${sheetRow}`,
          values: [[resolvedHome]]
        });
      }
      if (resolvedAway && resolvedAway !== m.away) {
        updates.push({
          range: `${SHEET_GM}!J${sheetRow}`,
          values: [[resolvedAway]]
        });
      }
    }

    if (!updates.length) return; // nothing to write

    // Batch write all updates in one API call
    await gapi.client.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      resource: {
        valueInputOption: 'RAW',
        data: updates
      }
    });

    console.log(`✅ Wrote ${updates.length} resolved KO team(s) to sheet`);
  } catch(e) {
    // Non-critical — don't surface to user, just log
    console.warn('writeResolvedKOTeams failed:', e);
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
        range: `${SHEET_GM}!F${sheetRow}:S${sheetRow}`
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
// ── Notifications ──────────────────────────────────────────
const _NOTIF_KEY = 'wc_notifications';

function getScheduledNotifs() {
  try { return JSON.parse(localStorage.getItem(_NOTIF_KEY) || '{}'); } catch(e) { return {}; }
}

function saveScheduledNotifs(obj) {
  localStorage.setItem(_NOTIF_KEY, JSON.stringify(obj));
}

function isNotifScheduled(matchNum) {
  return !!getScheduledNotifs()[matchNum];
}

async function registerSW() {
  if (!('serviceWorker' in navigator)) return null;
  try {
    const reg = await navigator.serviceWorker.register('/sw.js');
    return reg;
  } catch(e) { console.warn('SW register failed', e); return null; }
}

async function requestNotifPermission() {
  if (!('Notification' in window)) return false;
  if (Notification.permission === 'granted') return true;
  const result = await Notification.requestPermission();
  return result === 'granted';
}

function scheduleLocalNotif(title, body, fireAtMs, tag) {
  const delayMs = fireAtMs - Date.now();
  if (delayMs <= 0) return;
  setTimeout(() => {
    if (Notification.permission === 'granted') {
      new Notification(title, {
        body,
        icon: '/images/icon-192.png',
        tag,
      });
    }
  }, delayMs);
}

function matchKickoffMs(m) {
  // m.date = "11-Jun-26", m.time = "22:00" Cairo (UTC+3)
  try {
    const parts = m.date.split('-');
    const months = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
    const day   = parseInt(parts[0]);
    const month = months[parts[1]];
    const year  = 2000 + parseInt(parts[2]);
    const [hh, mm] = m.time.split(':').map(Number);
    // Cairo is UTC+3
    const utcMs = Date.UTC(year, month, day, hh - 3, mm);
    return utcMs;
  } catch(e) { return null; }
}

async function toggleMatchNotification(matchNum) {
  const matches = window._allMatches || [];
  const m = matches.find(x => x.num === matchNum);
  if (!m) return;

  const notifs = getScheduledNotifs();

  if (notifs[matchNum]) {
    // Cancel — just remove from storage (can't cancel setTimeout across sessions)
    delete notifs[matchNum];
    saveScheduledNotifs(notifs);
    showToast('🔕 Notification removed');
    refreshBellIcon(matchNum, false);
    return;
  }

  const granted = await requestNotifPermission();
  if (!granted) { showToast('Notifications blocked', 'error'); return; }

  const kickoffMs = matchKickoffMs(m);
  if (!kickoffMs) { showToast('Could not schedule', 'error'); return; }

  const beforeMs  = kickoffMs - (15 * 60 * 1000);
  const afterMs   = kickoffMs + (105 * 60 * 1000);
  const matchLabel = `${m.home} vs ${m.away}`;

  scheduleLocalNotif(
    '⚽ Match Starting Soon',
    `${matchLabel} kicks off in 15 minutes!`,
    beforeMs,
    `wc-before-${matchNum}`
  );
  scheduleLocalNotif(
    '🏁 Match Ended',
    `${matchLabel} has ended. Check the result!`,
    afterMs,
    `wc-after-${matchNum}`
  );

  notifs[matchNum] = { beforeMs, afterMs, label: matchLabel };
  saveScheduledNotifs(notifs);
  showToast('🔔 Notifications set!');
  refreshBellIcon(matchNum, true);
}

function refreshBellIcon(matchNum, active) {
  const btn = document.querySelector(`[data-notif-btn="${matchNum}"]`);
  if (btn) {
    btn.textContent = active ? '🔔' : '🔕';
    btn.style.opacity = active ? '1' : '0.5';
  }
}

function restoreScheduledNotifs() {
  const notifs = getScheduledNotifs();
  const now = Date.now();
  for (const [matchNum, n] of Object.entries(notifs)) {
    const m = (window._allMatches || []).find(x => x.num === parseInt(matchNum));
    if (!m) continue;
    const matchLabel = `${m.home} vs ${m.away}`;
    if (n.beforeMs > now) {
      scheduleLocalNotif('⚽ Match Starting Soon', `${matchLabel} kicks off in 15 minutes!`, n.beforeMs, `wc-before-${matchNum}`);
    }
    if (n.afterMs > now) {
      scheduleLocalNotif('🏁 Match Ended', `${matchLabel} has ended. Check the result!`, n.afterMs, `wc-after-${matchNum}`);
    }
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

/* ════════════════════════════════════════════════════════════
   R32 SEEDING — shared resolution logic
   Used by Seeding.html and Knockout.html
   ════════════════════════════════════════════════════════════ */

const R32_SEEDING = {
  73:  { home:{grp:'A', rank:'2'}, away:{grp:'B', rank:'2'} },
  74:  { home:{grp:'E', rank:'1'}, away:{grp:'A/B/C/D/F', rank:'3'} },
  75:  { home:{grp:'F', rank:'1'}, away:{grp:'C', rank:'2'} },
  76:  { home:{grp:'C', rank:'1'}, away:{grp:'F', rank:'2'} },
  77:  { home:{grp:'I', rank:'1'}, away:{grp:'C/D/F/G/H', rank:'3'} },
  78:  { home:{grp:'E', rank:'2'}, away:{grp:'I', rank:'2'} },
  79:  { home:{grp:'A', rank:'1'}, away:{grp:'C/E/F/H/I', rank:'3'} },
  80:  { home:{grp:'L', rank:'1'}, away:{grp:'E/H/I/J/K', rank:'3'} },
  81:  { home:{grp:'D', rank:'1'}, away:{grp:'B/E/F/I/J', rank:'3'} },
  82:  { home:{grp:'G', rank:'1'}, away:{grp:'A/E/H/I/J', rank:'3'} },
  83:  { home:{grp:'K', rank:'2'}, away:{grp:'L', rank:'2'} },
  84:  { home:{grp:'H', rank:'1'}, away:{grp:'J', rank:'2'} },
  85:  { home:{grp:'B', rank:'1'}, away:{grp:'E/F/G/I/J', rank:'3'} },
  86:  { home:{grp:'J', rank:'1'}, away:{grp:'H', rank:'2'} },
  87:  { home:{grp:'K', rank:'1'}, away:{grp:'D/E/I/J/L', rank:'3'} },
  88:  { home:{grp:'D', rank:'2'}, away:{grp:'G', rank:'2'} },
};

function computeGroupStandings(groupId, allMatches) {
  const teams = GROUPS_TEAMS[groupId] || [];
  const stats = {};
  teams.forEach(t => { stats[t] = { played:0, wins:0, draws:0, losses:0, gf:0, ga:0, gd:0, points:0, fairPlay:0 }; });

  allMatches.filter(m => m.group === groupId).forEach(m => {
    if (m.status !== 'Played') return;
    const h = m.home, a = m.away;
    if (!stats[h] || !stats[a]) return;
    const hg = parseInt(m.homeScore)||0, ag = parseInt(m.awayScore)||0;
    stats[h].played++; stats[a].played++;
    stats[h].gf += hg; stats[h].ga += ag;
    stats[a].gf += ag; stats[a].ga += hg;
    if (hg > ag)      { stats[h].wins++; stats[a].losses++; stats[h].points += 3; }
    else if (ag > hg) { stats[a].wins++; stats[h].losses++; stats[a].points += 3; }
    else              { stats[h].draws++; stats[a].draws++; stats[h].points++; stats[a].points++; }
    stats[h].fairPlay += (parseInt(m.homeYellow)||0)*-1 + (parseInt(m.homeRed)||0)*-3;
    stats[a].fairPlay += (parseInt(m.awayYellow)||0)*-1 + (parseInt(m.awayRed)||0)*-3;
  });
  for (let t in stats) stats[t].gd = stats[t].gf - stats[t].ga;

  function getH2H(tied) {
    const h2h = {};
    tied.forEach(t => { h2h[t] = { points:0, gd:0, gf:0 }; });
    allMatches.filter(m => m.group === groupId && m.status === 'Played'
      && tied.includes(m.home) && tied.includes(m.away)).forEach(m => {
      const hg = parseInt(m.homeScore)||0, ag = parseInt(m.awayScore)||0;
      h2h[m.home].gf += hg; h2h[m.home].gd += hg - ag;
      h2h[m.away].gf += ag; h2h[m.away].gd += ag - hg;
      if (hg > ag) h2h[m.home].points += 3;
      else if (ag > hg) h2h[m.away].points += 3;
      else { h2h[m.home].points++; h2h[m.away].points++; }
    });
    return h2h;
  }

  let standings = teams.map(t => ({ team:t, ...stats[t] }));
  standings.sort((a, b) => {
    if (a.points !== b.points) return b.points - a.points;
    if (a.gd !== b.gd) return b.gd - a.gd;
    if (a.gf !== b.gf) return b.gf - a.gf;
    const tied = standings.filter(t =>
      t.points === a.points && t.gd === a.gd && t.gf === a.gf).map(t => t.team);
    if (tied.length === 2) {
      const h2h = getH2H(tied);
      if (h2h[a.team].points !== h2h[b.team].points) return h2h[b.team].points - h2h[a.team].points;
      if (h2h[a.team].gd    !== h2h[b.team].gd)     return h2h[b.team].gd     - h2h[a.team].gd;
      if (h2h[a.team].gf    !== h2h[b.team].gf)     return h2h[b.team].gf     - h2h[a.team].gf;
    }
    if (a.fairPlay !== b.fairPlay) return b.fairPlay - a.fairPlay;
    return 0;
  });
  return standings;
}

function isGroupComplete(groupId, allMatches) {
  const gm = allMatches.filter(m => m.group === groupId);
  const played = gm.filter(m => m.status === 'Played').length;
  return gm.length > 0 && played >= 6;
}

function buildAllStandings(allMatches) {
  const standings = {};
  const thirdPlaceTeams = {};
  for (const grp of ['A','B','C','D','E','F','G','H','I','J','K','L']) {
    standings[grp] = computeGroupStandings(grp, allMatches);
    if (standings[grp][2]) thirdPlaceTeams[grp] = standings[grp][2];
  }
  return { standings, thirdPlaceTeams };
}

function resolveBestThird(eligibleStr, thirdPlaceTeams, allMatches) {
  const eligible = eligibleStr.split('/');
  if (!eligible.every(g => isGroupComplete(g, allMatches))) return null;
  let best = null;
  for (const grp of eligible) {
    const t = thirdPlaceTeams[grp];
    if (!t) continue;
    if (!best) { best = t; continue; }
    if (t.points > best.points) { best = t; continue; }
    if (t.points === best.points) {
      if (t.gd > best.gd) { best = t; continue; }
      if (t.gd === best.gd && t.gf > best.gf) { best = t; continue; }
    }
  }
  return best ? best.team : null;
}

/* Used by Knockout.html — only shows team when group is fully complete (all 6 matches played) */
function resolveR32Slot(slot, standings, thirdPlaceTeams, allMatches) {
  if (slot.rank === '3') {
    return resolveBestThird(slot.grp, thirdPlaceTeams, allMatches);
  }
  if (!isGroupComplete(slot.grp, allMatches)) return null;
  const st = standings[slot.grp];
  if (!st) return null;
  return st[slot.rank === '1' ? 0 : 1]?.team || null;
}

/* Used by Seeding.html — shows current standings leader dynamically after each match */
function resolveR32SlotDynamic(slot, standings, thirdPlaceTeams, allMatches) {
  if (slot.rank === '3') {
    return resolveBestThird(slot.grp, thirdPlaceTeams, allMatches);
  }
  const st = standings[slot.grp];
  if (!st) return null;
  const idx = slot.rank === '1' ? 0 : 1;
  const team = st[idx];
  if (!team || team.played === 0) return null;
  return team.team;
}
