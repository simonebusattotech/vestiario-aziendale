/**
 * Vestiario aziendale - backend su Google Sheets (Google Apps Script)
 *
 * Installazione (una volta sola):
 *  1. Crea un nuovo Foglio Google (es. "Vestiario").
 *  2. Menu Estensioni > Apps Script, cancella il contenuto e incolla questo file.
 *  3. Cambia API_KEY qui sotto con una parola segreta a tua scelta.
 *  4. Seleziona la funzione "setup" e premi Esegui (autorizza l'accesso quando richiesto).
 *  5. Distribuisci > Nuovo deployment > tipo "App web"
 *       Esegui come: Me    |    Chi ha accesso: Chiunque
 *     Copia l'URL che termina con /exec.
 *
 * Il foglio contiene:
 *  - Capi       : un capo per riga con il CID del proprietario e lo stato
 *  - Movimenti  : storico di ogni lettura
 *  - Utenti     : (facoltativo) elenco dei CID ammessi. Se vuoto, qualsiasi CID è accettato.
 */

const API_KEY = 'CAMBIA-QUESTA-CHIAVE';

const SH_CAPI = 'Capi';
const SH_MOV = 'Movimenti';
const SH_UTENTI = 'Utenti';
const CAPI_COLS = ['Codice', 'CID', 'Stato', 'Data registrazione', 'Ultimo invio', 'Ultimo rientro', 'Nome'];
const MOV_COLS = ['Data/Ora', 'CID', 'Codice', 'Operazione', 'Esito', 'Note', 'Nome'];
const UTENTI_COLS = ['CID', 'Nome'];

const LAV = 'In lavanderia';
const SEDE = 'In sede';
const MODES = { REG: 'REGISTRAZIONE', INVIO: 'INVIO LAVANDERIA', RIENTRO: 'RIENTRO / VERIFICA' };

// --------------------------------------------------------------------------
// Setup
// --------------------------------------------------------------------------
function setup() {
  const capi = getSheet_(SH_CAPI, CAPI_COLS);
  const mov = getSheet_(SH_MOV, MOV_COLS);
  const ut = getSheet_(SH_UTENTI, UTENTI_COLS);

  // Codici e CID come TESTO (evita la perdita degli zeri iniziali nei barcode numerici)
  capi.getRange('A:B').setNumberFormat('@');
  mov.getRange('B:C').setNumberFormat('@');
  ut.getRange('A:A').setNumberFormat('@');
  capi.getRange('D:F').setNumberFormat('dd/mm/yyyy hh:mm');
  mov.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm:ss');
  [[capi, [160, 100, 120, 150, 150, 150, 200]], [mov, [150, 100, 160, 160, 100, 320, 200]], [ut, [100, 220]]]
    .forEach(([sh, w]) => w.forEach((px, i) => sh.setColumnWidth(i + 1, px)));

  // Colori automatici
  const rule = (range, text, color) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text).setBackground(color).setRanges([range]).build();
  capi.setConditionalFormatRules([
    rule(capi.getRange('C2:C'), LAV, '#fff2cc'),
    rule(capi.getRange('C2:C'), SEDE, '#d9ead3'),
  ]);
  mov.setConditionalFormatRules([
    rule(mov.getRange('E2:E'), 'ERRORE', '#f4cccc'),
    rule(mov.getRange('E2:E'), 'ATTENZIONE', '#fce5cd'),
    rule(mov.getRange('E2:E'), 'OK', '#d9ead3'),
  ]);

  const def = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Foglio1') ||
              SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(def);

  aggiornaNomi();
}

function getSheet_(name, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.setFrozenRows(1);
  }
  // crea o completa l'intestazione (es. colonna "Nome" aggiunta in seguito)
  if (sh.getRange(1, cols.length).getValue() !== cols[cols.length - 1]) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols])
      .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  }
  return sh;
}

// --------------------------------------------------------------------------
// Nomi utente
// --------------------------------------------------------------------------
/** Mappa CID -> Nome dalla scheda Utenti. */
function nomiUtenti_() {
  const map = {};
  readRows_(getSheet_(SH_UTENTI, UTENTI_COLS), UTENTI_COLS.length)
    .forEach(r => { if (norm_(r[0])) map[norm_(r[0])] = String(r[1] || '').trim(); });
  return map;
}

/** Riscrive la colonna Nome della scheda Capi in base alla scheda Utenti. */
function aggiornaNomi() {
  const sh = getSheet_(SH_CAPI, CAPI_COLS);
  const rows = readRows_(sh, CAPI_COLS.length);
  if (!rows.length) return;
  const nomi = nomiUtenti_();
  const col = CAPI_COLS.indexOf('Nome') + 1;
  sh.getRange(2, col, rows.length, 1).setValues(rows.map(r => [nomi[norm_(r[1])] || '']));
}

/** Trigger automatico: quando modifichi la scheda Utenti, i nomi nei Capi si aggiornano. */
function onEdit(e) {
  if (e && e.range && e.range.getSheet().getName() === SH_UTENTI) aggiornaNomi();
}

/** Menu nel foglio: Vestiario > Aggiorna nomi. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Vestiario')
    .addItem('Aggiorna nomi da Utenti', 'aggiornaNomi')
    .addToUi();
}

// --------------------------------------------------------------------------
// API web
// --------------------------------------------------------------------------
function doGet() {
  return json_({ ok: true, msg: 'API Vestiario attiva' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Richiesta non valida' });
  }
  if (req.key !== API_KEY) return json_({ ok: false, error: 'Chiave di accesso non valida' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, error: 'Server occupato, riprova' });
  try {
    const cid = norm_(req.cid);
    if (!cid) throw new Error('CID mancante');
    const nome = checkCid_(cid);
    switch (req.action) {
      case 'login':
        return json_({ ok: true, cid: cid, nome: nome, capi: capiDi_(cid) });
      case 'scan':
        return json_(Object.assign({ ok: true }, scan_(cid, req.mode, norm_(req.code), nome)));
      case 'chiudi':
        return json_({ ok: true, mancanti: chiudi_(cid, nome) });
      default:
        throw new Error('Azione sconosciuta');
    }
  } catch (err) {
    return json_({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// --------------------------------------------------------------------------
// Logica
// --------------------------------------------------------------------------
function scan_(cid, mode, code, nome) {
  if (!code) throw new Error('Codice vuoto');
  const sh = getSheet_(SH_CAPI, CAPI_COLS);
  const rows = readRows_(sh, CAPI_COLS.length);
  const i = rows.findIndex(r => norm_(r[0]) === code);
  const rec = i >= 0 ? rows[i] : null;
  const owner = rec ? norm_(rec[1]) : null;
  const ownerTxt = owner && owner !== cid ? owner + (n => n ? ' (' + n + ')' : '')(nomiUtenti_()[owner]) : '';
  const now = new Date();
  const save = r => sh.getRange(i + 2, 1, 1, CAPI_COLS.length).setValues([r]);
  let esito, msg;

  if (mode === MODES.REG) {
    if (!rec) {
      sh.appendRow([code, cid, SEDE, now, '', '', nome]);
      esito = 'OK'; msg = 'Capo registrato';
    } else if (owner === cid) {
      esito = 'ATTENZIONE'; msg = 'Capo già registrato a questo CID';
    } else {
      esito = 'ERRORE'; msg = 'Capo già registrato al CID ' + ownerTxt;
    }
  } else if (mode === MODES.INVIO) {
    if (!rec) {
      sh.appendRow([code, cid, LAV, now, now, '', nome]);
      esito = 'OK'; msg = 'Nuovo capo registrato e inviato in lavanderia';
    } else if (owner !== cid) {
      esito = 'ERRORE'; msg = 'Questo capo appartiene al CID ' + ownerTxt;
    } else if (rec[2] === LAV) {
      esito = 'ATTENZIONE'; msg = 'Capo già segnato in lavanderia';
    } else {
      rec[2] = LAV; rec[4] = now; rec[6] = nome; save(rec);
      esito = 'OK'; msg = 'Inviato in lavanderia';
    }
  } else if (mode === MODES.RIENTRO) {
    if (!rec) {
      // capo mai visto prima: lo registra al CID, ma lo segnala in arancione
      sh.appendRow([code, cid, SEDE, now, '', now, nome]);
      esito = 'ATTENZIONE'; msg = 'Capo nuovo (mai inviato): registrato al tuo CID';
    } else if (owner !== cid) {
      esito = 'ERRORE'; msg = 'CAPO NON TUO - appartiene al CID ' + ownerTxt;
    } else if (rec[2] !== LAV) {
      esito = 'ATTENZIONE'; msg = 'Capo già rientrato / non risultava in lavanderia';
    } else {
      rec[2] = SEDE; rec[5] = now; rec[6] = nome; save(rec);
      esito = 'OK'; msg = 'Capo corretto - rientrato';
    }
  } else {
    throw new Error('Operazione non valida');
  }

  log_([[now, cid, code, mode, esito, msg, nome]]);
  return { esito: esito, msg: msg, code: code, capi: capiDi_(cid) };
}

function chiudi_(cid, nome) {
  const mancanti = capiDi_(cid).filter(c => c.stato === LAV).map(c => c.codice);
  if (mancanti.length) {
    const now = new Date();
    log_(mancanti.map(c => [now, cid, c, MODES.RIENTRO, 'ATTENZIONE', 'Capo NON rientrato dalla lavanderia', nome]));
  }
  return mancanti;
}

function capiDi_(cid) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const fmt = d => d instanceof Date ? Utilities.formatDate(d, tz, 'dd/MM/yyyy HH:mm') : String(d || '');
  return readRows_(getSheet_(SH_CAPI, CAPI_COLS), CAPI_COLS.length)
    .filter(r => norm_(r[1]) === cid)
    .map(r => ({ codice: norm_(r[0]), stato: r[2], invio: fmt(r[4]), rientro: fmt(r[5]) }))
    .sort((a, b) => a.codice.localeCompare(b.codice));
}

function checkCid_(cid) {
  const sh = getSheet_(SH_UTENTI, UTENTI_COLS);
  const rows = readRows_(sh, UTENTI_COLS.length).filter(r => norm_(r[0]));
  if (!rows.length) return '';               // nessun elenco: tutti i CID ammessi
  const u = rows.find(r => norm_(r[0]) === cid);
  if (!u) throw new Error('CID ' + cid + ' non abilitato');
  return String(u[1] || '');
}

function log_(rows) {
  const sh = getSheet_(SH_MOV, MOV_COLS);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, MOV_COLS.length).setValues(rows);
}

function readRows_(sh, ncol) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, ncol).getValues() : [];
}

function norm_(v) {
  return String(v == null ? '' : v).trim().toUpperCase();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
