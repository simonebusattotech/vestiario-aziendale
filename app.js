'use strict';

const MODES = {
  'REGISTRAZIONE': 'Associa i capi nuovi al tuo CID.',
  'INVIO LAVANDERIA': 'Scansiona i capi che consegni alla lavanderia.',
  'RIENTRO / VERIFICA': 'Scansiona i capi tornati: verifichiamo che siano i tuoi e che non ne manchi nessuno.',
};
const LAV = 'In lavanderia';

const $ = id => document.getElementById(id);
const IS_TOUCH = window.matchMedia('(pointer: coarse)').matches;
const state = {
  mode: 'INVIO LAVANDERIA',
  cid: null,
  capi: [],
  letti: new Set(),
  busy: false,
  scanner: null,
  lastCode: '',
  lastTime: 0,
};

// ------------------------------------------------------------------ config
const store = {
  get(k) { try { return localStorage.getItem(k) || ''; } catch { return ''; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch { /* storage non disponibile */ } },
};
// URL e chiave NON sono nel codice (repository pubblico): arrivano dal link di configurazione
// (?api=…&key=…) o da ⚙ Impostazioni e restano salvati solo sul dispositivo.
const cfg = () => ({ url: store.get('apiUrl'), key: store.get('apiKey') });

// Configurazione tramite link:  index.html?api=URL&key=CHIAVE
(function readSetupLink() {
  const p = new URLSearchParams(location.search);
  if (p.get('api')) {
    store.set('apiUrl', p.get('api'));
    if (p.get('key')) store.set('apiKey', p.get('key'));
    history.replaceState(null, '', location.pathname);
  }
})();

async function api(action, data = {}) {
  const { url, key } = cfg();
  if (!url) throw new Error('App non configurata: apri ⚙ Impostazioni');
  if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(url)) {
    throw new Error('URL non valido: deve essere l\'URL dell\'app web che termina con /exec (⚙ Impostazioni)');
  }
  let res;
  try {
    res = await fetch(url, {
      method: 'POST',
      body: JSON.stringify({ action, key, ...data }), // text/plain: niente preflight CORS
    });
  } catch {
    throw new Error(navigator.onLine === false
      ? 'Nessuna connessione internet'
      : 'Lo script Google non risponde. Controlla che il deployment sia "App web" con accesso "Chiunque"');
  }
  if (!res.ok) throw new Error('Errore server (' + res.status + ')');
  let json;
  try {
    json = await res.json();
  } catch {
    throw new Error('Risposta non valida dallo script: controlla il deployment e l\'autorizzazione');
  }
  if (!json.ok) throw new Error(json.error || 'Errore sconosciuto');
  return json;
}

// ------------------------------------------------------------------ feedback
let audioCtx;
function beep(ok) {
  try {
    audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
    const o = audioCtx.createOscillator(), g = audioCtx.createGain();
    o.frequency.value = ok ? 1200 : 300;
    o.type = ok ? 'sine' : 'square';
    g.gain.value = 0.15;
    o.connect(g).connect(audioCtx.destination);
    o.start();
    o.stop(audioCtx.currentTime + (ok ? 0.12 : 0.5));
  } catch { /* audio non disponibile */ }
  if (navigator.vibrate) navigator.vibrate(ok ? 80 : [200, 80, 200]);
}

function banner(text, kind = 'idle', code = '') {
  const b = $('banner');
  b.className = 'banner ' + kind;
  b.textContent = '';
  if (code) {
    const s = document.createElement('span');
    s.className = 'code';
    s.textContent = code;
    b.appendChild(s);
  }
  b.appendChild(document.createTextNode(text));
}

const kindOf = esito => ({ OK: 'ok', ATTENZIONE: 'warn' }[esito] || 'err');

// ------------------------------------------------------------------ render
function renderMode() {
  document.querySelectorAll('.modes button').forEach(b =>
    b.classList.toggle('active', b.dataset.mode === state.mode));
  $('modeHint').textContent = MODES[state.mode];
}

function renderCapi() {
  const ul = $('capiList');
  ul.textContent = '';
  if (!state.capi.length) {
    const li = document.createElement('li');
    li.className = 'empty';
    li.textContent = 'Nessun capo registrato';
    ul.appendChild(li);
  }
  for (const c of state.capi) {
    const li = document.createElement('li');
    li.className = state.letti.has(c.codice) ? 'letto' : (c.stato === LAV ? 'lav' : 'sede');
    const code = document.createElement('span');
    code.className = 'c';
    code.textContent = c.codice;
    const st = document.createElement('span');
    st.className = 's';
    st.textContent = c.stato + (c.stato === LAV && c.invio ? '\n' + c.invio : '');
    st.style.whiteSpace = 'pre-line';
    li.append(code, st);
    ul.appendChild(li);
  }

  const counter = $('counter');
  const inLav = state.capi.filter(c => c.stato === LAV).length;
  if (state.mode === 'RIENTRO / VERIFICA') {
    counter.textContent = `Rientrati ${state.letti.size} · Mancanti ${inLav}`;
    counter.className = 'counter ' + (inLav ? 'warn' : 'ok');
  } else if (state.mode === 'INVIO LAVANDERIA') {
    counter.textContent = `In lavanderia ${inLav}/${state.capi.length}`;
    counter.className = 'counter';
  } else {
    counter.textContent = `Registrati ${state.capi.length}`;
    counter.className = 'counter';
  }
}

function addLog(code, esito, msg) {
  const li = document.createElement('li');
  const t = document.createElement('span');
  t.className = 't';
  t.textContent = new Date().toLocaleTimeString('it-IT') + '  ';
  const e = document.createElement('span');
  e.className = esito;
  e.textContent = esito + ' ';
  const c = document.createElement('b');
  c.textContent = code;
  li.append(t, e, c, document.createTextNode(' — ' + msg));
  $('logList').prepend(li);
}

function showSession(active) {
  $('loginCard').classList.toggle('hidden', active);
  $('sessionCard').classList.toggle('hidden', !active);
  $('listCard').classList.toggle('hidden', !active);
  $('logCard').classList.toggle('hidden', !active);
}

// ------------------------------------------------------------------ sessione
async function login() {
  const cid = $('cidInput').value.trim().toUpperCase();
  if (!cid) { banner('Inserisci il tuo CID', 'err'); beep(false); return; }
  setBusy(true, 'Verifica CID…');
  try {
    const r = await api('login', { cid });
    state.cid = r.cid;
    state.capi = r.capi;
    state.letti = new Set();
    $('cidLabel').textContent = r.cid + (r.nome ? ' · ' + r.nome : '');
    $('logList').textContent = '';
    showSession(true);
    renderCapi();
    if (state.mode === 'RIENTRO / VERIFICA' && !state.capi.some(c => c.stato === LAV)) {
      banner('Nessun tuo capo risulta in lavanderia', 'warn');
    } else {
      banner('Scansiona i capi', 'idle');
    }
    // Telefono/tablet: fotocamera. PC: pronto per il lettore USB (nessuna fotocamera automatica).
    if (IS_TOUCH) startCamera();
    else $('codeInput').focus();
  } catch (err) {
    banner(err.message, 'err'); beep(false);
  } finally {
    setBusy(false);
  }
}

async function endSession() {
  if (state.busy) return false;
  if (state.mode === 'RIENTRO / VERIFICA') {
    const mancanti = state.capi.filter(c => c.stato === LAV).map(c => c.codice);
    if (mancanti.length) {
      const ok = confirm(`Mancano ${mancanti.length} capi non rientrati:\n\n` +
        mancanti.slice(0, 20).join('\n') + (mancanti.length > 20 ? '\n…' : '') +
        '\n\nRegistrare la mancanza e terminare?');
      if (!ok) return false;
      setBusy(true, 'Salvataggio…');
      try {
        await api('chiudi', { cid: state.cid });
      } catch (err) {
        banner(err.message, 'err'); beep(false); setBusy(false);
        return false;
      }
      setBusy(false);
      banner(`Verifica chiusa: ${mancanti.length} capi mancanti segnalati`, 'warn');
    } else {
      banner('Verifica completata: tutti i capi sono rientrati', 'ok');
    }
  } else {
    banner('Sessione terminata. Inserisci il CID per iniziare', 'idle');
  }
  await stopCamera();
  state.cid = null;
  state.capi = [];
  state.letti = new Set();
  $('cidInput').value = '';
  showSession(false);
  return true;
}

async function handleCode(raw) {
  const code = String(raw || '').trim().toUpperCase();
  if (!code || !state.cid || state.busy) return;
  const now = Date.now();
  if (code === state.lastCode && now - state.lastTime < 3000) return; // evita doppie letture
  state.lastCode = code;
  state.lastTime = now;

  setBusy(true, 'Invio in corso…', code);
  try {
    const r = await api('scan', { cid: state.cid, mode: state.mode, code });
    state.capi = r.capi;
    if (state.mode === 'RIENTRO / VERIFICA' && r.esito === 'OK') state.letti.add(r.code);
    banner(r.msg, kindOf(r.esito), r.code);
    addLog(r.code, r.esito, r.msg);
    beep(r.esito === 'OK');
    renderCapi();
  } catch (err) {
    banner(err.message, 'err', code);
    addLog(code, 'ERRORE', err.message + ' (non salvato)');
    beep(false);
  } finally {
    setBusy(false);
  }
}

function setBusy(busy, text, code) {
  state.busy = busy;
  ['btnLogin', 'btnEnd', 'btnCode'].forEach(id => { $(id).disabled = busy; });
  document.querySelectorAll('.modes button').forEach(b => { b.disabled = busy; });
  if (busy && text) banner(text, 'busy', code);
}

// ------------------------------------------------------------------ fotocamera
async function startCamera() {
  if (state.scanner || typeof Html5Qrcode === 'undefined') {
    if (typeof Html5Qrcode === 'undefined') banner('Lettore fotocamera non caricato: usa il campo manuale', 'warn');
    return;
  }
  const F = Html5QrcodeSupportedFormats;
  state.scanner = new Html5Qrcode('reader', {
    verbose: false,
    formatsToSupport: [F.CODE_128, F.CODE_39, F.CODE_93, F.EAN_13, F.EAN_8, F.UPC_A, F.UPC_E,
                       F.ITF, F.CODABAR, F.QR_CODE, F.DATA_MATRIX],
    experimentalFeatures: { useBarCodeDetectorIfSupported: true },
  });
  try {
    await state.scanner.start(
      { facingMode: 'environment' },
      { fps: 10, qrbox: (w, h) => ({ width: Math.max(60, Math.floor(w * 0.9)), height: Math.max(60, Math.floor(Math.min(w, h) * 0.45)) }) },
      text => handleCode(text),
      () => {},
    );
    $('btnCamera').textContent = '⏹ Ferma fotocamera';
  } catch (err) {
    state.scanner = null;
    $('reader').textContent = '';
    banner('Fotocamera non disponibile: consenti l\'accesso o usa il campo manuale', 'warn');
  }
}

async function stopCamera() {
  if (!state.scanner) return;
  try { await state.scanner.stop(); state.scanner.clear(); } catch { /* già fermo */ }
  state.scanner = null;
  $('reader').textContent = '';
  $('btnCamera').textContent = '📷 Avvia fotocamera';
}

// ------------------------------------------------------------------ eventi
document.querySelectorAll('.modes button').forEach(btn => btn.addEventListener('click', async () => {
  const mode = btn.dataset.mode;
  if (mode === state.mode) return;
  if (state.cid && !(await endSession())) return;
  state.mode = mode;
  renderMode();
}));

$('btnLogin').addEventListener('click', login);
$('cidInput').addEventListener('keydown', e => { if (e.key === 'Enter') login(); });
$('btnEnd').addEventListener('click', endSession);
$('btnCamera').addEventListener('click', () => (state.scanner ? stopCamera() : startCamera()));

function submitManual() {
  const v = $('codeInput').value;
  $('codeInput').value = '';
  handleCode(v);
  $('codeInput').focus();
}
$('btnCode').addEventListener('click', submitManual);
$('codeInput').addEventListener('keydown', e => { if (e.key === 'Enter') submitManual(); });

// Lettore barcode USB/bluetooth (emula una tastiera): cattura i caratteri anche se nessun
// campo è selezionato, così basta "sparare" il codice senza toccare lo schermo.
let scanBuf = '', scanTimer = null;
document.addEventListener('keydown', e => {
  if (!state.cid || e.ctrlKey || e.altKey || e.metaKey) return;
  const el = document.activeElement;
  if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') || $('settings').open) return;
  if (e.key === 'Enter') {
    if (scanBuf) { e.preventDefault(); handleCode(scanBuf); }
    scanBuf = '';
  } else if (e.key.length === 1) {
    scanBuf += e.key;
    clearTimeout(scanTimer);
    scanTimer = setTimeout(() => { scanBuf = ''; }, 500); // un lettore invia tutto in pochi ms
  }
});

// impostazioni
$('btnSettings').addEventListener('click', () => {
  const { url, key } = cfg();
  $('apiUrl').value = url;
  $('apiKey').value = key;
  $('settings').showModal();
});
$('btnSave').addEventListener('click', () => {
  store.set('apiUrl', $('apiUrl').value.trim());
  store.set('apiKey', $('apiKey').value.trim());
  banner('Impostazioni salvate', 'ok');
});
$('btnCopyLink').addEventListener('click', async () => {
  const link = location.origin + location.pathname + '?api=' + encodeURIComponent($('apiUrl').value.trim()) +
               '&key=' + encodeURIComponent($('apiKey').value.trim());
  try {
    await navigator.clipboard.writeText(link);
    $('btnCopyLink').textContent = 'Link copiato ✓';
  } catch {
    prompt('Copia questo link e invialo ai colleghi:', link);
  }
});

renderMode();
if (!cfg().url || !cfg().key) {
  banner('App non configurata: apri il link di configurazione ricevuto oppure ⚙ Impostazioni', 'warn');
}
