/**
 * Fioriture Civiche — Raccontaci il tuo quartiere
 * Backend minimo: Google Apps Script Web App che riceve le risposte (POST JSON)
 * e le scrive su un foglio Google. Le foto vengono salvate in una cartella Drive.
 *
 * SETUP (≈10 minuti):
 *  1. Crea un foglio Google vuoto (es. "Fioriture Civiche – Risposte questionario").
 *  2. Estensioni → Apps Script. Incolla questo file al posto di Code.gs. Salva.
 *  3. (facoltativo) Crea una cartella Drive per le foto e incolla il suo ID in PHOTO_FOLDER_ID.
 *     Se resta vuoto, la cartella "Fioriture Civiche – Foto questionario" viene creata al primo invio.
 *  4. Distribuisci → Nuova distribuzione → tipo "Applicazione web":
 *       Esegui come: Me        |  Chi può accedere: Chiunque
 *     → Distribuisci → autorizza → copia l'URL (termina con /exec).
 *  5. In index.html imposta  const ENDPOINT = '<URL copiato>';  e fai il push.
 *  Ogni modifica a questo script richiede "Gestisci distribuzioni → modifica → nuova versione".
 */

const PHOTO_FOLDER_ID = ''; // opzionale

const COLS = [
  'id','ricevuto','inizio','durata_s','consenso','lat','lng',
  'A1','A1_foto','A2','A3','A3_foto',
  'B1_n',
  'C1','C1b','C1b_altro','C2','C3','C4_n','C5a_n','C5b_n','C6_giorno','C6_buio',
  'D1',
  'eta','genere','vive','giorno','scuola_nome','attivita','attivita_altro',
  'ua','json'
];
const PIN_COLS = ['id_risposta','domanda','n','lat','lng','tipo','fermata','perche'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const id = Utilities.getUuid().slice(0, 8);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const r = data.risposte || {};
    const a = data.auto || {};

    // foto → Drive
    const fotoUrl = {};
    ['A1', 'A3'].forEach(k => {
      const img = r[k + '_foto'];
      if (img && typeof img === 'string' && img.startsWith('data:image')) {
        fotoUrl[k] = savePhoto_(img, id + '_' + k);
      }
    });

    // riga risposte
    const sh = sheet_(ss, 'risposte', COLS);
    const row = {
      id, ricevuto: new Date(), inizio: a.inizio || '', durata_s: a.durata_s || '', consenso: a.consenso ? 'sì' : 'no',
      lat: a.gps ? a.gps[0] : '', lng: a.gps ? a.gps[1] : '',
      A1: r.A1 || '', A1_foto: fotoUrl.A1 || '', A2: r.A2 || '', A3: r.A3 || '', A3_foto: fotoUrl.A3 || '',
      B1_n: (r.B1 || []).length,
      C1: r.C1 || '', C1b: join_(r.C1b), C1b_altro: r.C1b_altro || '', C2: join_(r.C2), C3: r.C3 || '',
      C4_n: (r.C4 || []).length, C5a_n: (r.C5a || []).length, C5b_n: (r.C5b || []).length,
      C6_giorno: r.C6_giorno || '', C6_buio: r.C6_buio || '',
      D1: r.D1 || '',
      eta: r.eta || '', genere: r.genere || '', vive: r.vive || '', giorno: r.giorno || '', scuola_nome: r.scuola_nome || '',
      attivita: join_(r.attivita), attivita_altro: r.attivita_altro || '',
      ua: a.ua || '', json: JSON.stringify(stripPhotos_(data))
    };
    sh.appendRow(COLS.map(c => row[c] === undefined ? '' : row[c]));

    // righe luoghi (una per pin)
    const shp = sheet_(ss, 'luoghi', PIN_COLS);
    ['B1', 'C4', 'C5a', 'C5b'].forEach(q => (r[q] || []).forEach((p, i) => {
      shp.appendRow([id, q, i + 1, p.lat, p.lng, p.tipo || '', p.fermata || '', p.perche || '']);
    }));

    return out_({ ok: true, id });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doGet() { return out_({ ok: true, service: 'fioriture-civiche-questionario' }); }

// ---- helpers ----
function sheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(cols); sh.setFrozenRows(1); }
  return sh;
}
function join_(v) { return Array.isArray(v) ? v.join(' | ') : (v || ''); }
function stripPhotos_(d) {
  const c = JSON.parse(JSON.stringify(d));
  Object.keys(c.risposte || {}).forEach(k => { if (typeof c.risposte[k] === 'string' && c.risposte[k].startsWith('data:image')) c.risposte[k] = '[foto]'; });
  return c;
}
function savePhoto_(dataUrl, name) {
  const m = dataUrl.match(/^data:(image\/\w+);base64,(.*)$/);
  if (!m) return '';
  const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name + '.' + ext);
  const folder = PHOTO_FOLDER_ID ? DriveApp.getFolderById(PHOTO_FOLDER_ID) : folderByName_('Fioriture Civiche – Foto questionario');
  return folder.createFile(blob).getUrl();
}
function folderByName_(n) {
  const it = DriveApp.getFoldersByName(n);
  return it.hasNext() ? it.next() : DriveApp.createFolder(n);
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
