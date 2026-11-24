/**
 * ORDINI LIBRI DI NATALE — Antonio Catalano / Casa degli Alfieri
 * Script per il foglio Google "Ordini libri di Natale".
 *
 * Cosa fa:
 *  1. Riceve ogni ordine pagato dal sito e lo scrive subito nel foglio "Ordini".
 *  2. Ogni 3 ore: se ci sono ordini nuovi, manda UNA mail con l'elenco. Se non ce ne sono, non manda niente.
 *  3. Ogni 10 ore: se sono arrivati nuovi ordini con ritiro a Bergamo, manda il riepilogo per la
 *     biglietteria del teatro (tutti i ritiri ancora da fare). Se non ce ne sono, non manda niente.
 *  Tutte le mail vanno a DESTINATARIO.
 *
 * Installazione: vedi le istruzioni in fondo a questo file.
 */

const DESTINATARIO = 'universisensibili@sensibili.it';
const CHIAVE = 'cambia-questa-parola-segreta';   // la stessa va in NEGOZIO.ordiniChiave nel sito
const ORE_NOTIFICA_ORDINI = 3;
const ORE_RIEPILOGO_BIGLIETTERIA = 10;
const FOGLIO = 'Ordini';

const COLONNE = ['Data', 'Codice', 'Nome', 'Email', 'Telefono', 'Consegna', 'Indirizzo',
  'Libri', 'Copie', 'Totale €', 'Pagamento', 'Note', 'Stato', 'Notificato il', 'In riepilogo il'];
const C = {}; COLONNE.forEach((n, i) => C[n] = i + 1);
const STATI = ['Da ritirare', 'Ritirato', 'Da spedire', 'Spedito', 'Annullato'];

/* ---------- 1. ricezione ordini dal sito ---------- */
function doPost(e) {
  let o;
  try { o = JSON.parse(e.postData.contents); } catch (err) { return risposta({ ok: false, errore: 'json' }); }
  if (o.chiave !== CHIAVE) return risposta({ ok: false, errore: 'chiave' });
  if (!o.codice || !o.nome || !o.email || !Array.isArray(o.righe) || !o.righe.length) return risposta({ ok: false, errore: 'dati' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = foglio();
    const codici = sh.getLastRow() > 1 ? sh.getRange(2, C['Codice'], sh.getLastRow() - 1, 1).getValues().flat() : [];
    if (codici.indexOf(o.codice) >= 0) return risposta({ ok: true, doppio: true });   // già registrato

    const ritiro = String(o.consegna).indexOf('Ritiro') === 0;
    sh.appendRow([
      new Date(o.data || Date.now()), o.codice, o.nome, o.email, o.telefono || '',
      ritiro ? 'Ritiro a Bergamo' : 'Spedizione', o.indirizzo || '',
      o.righe.map(r => `${r.copie} × ${r.titolo}`).join('\n'), Number(o.copie) || 0, Number(o.totale) || 0,
      o.pagamento || '', o.note || '', ritiro ? 'Da ritirare' : 'Da spedire', '', ''
    ]);
  } finally { lock.releaseLock(); }
  return risposta({ ok: true });
}
function risposta(x) { return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- 2 e 3. controllo orario (gira ogni ora, decide lui se è il momento) ---------- */
function controlloOrario() {
  const p = PropertiesService.getScriptProperties(), ora = Date.now();
  if (ora - Number(p.getProperty('ultimaNotifica') || 0) >= ORE_NOTIFICA_ORDINI * 3600e3 - 5 * 60e3) {
    notificaOrdini(); p.setProperty('ultimaNotifica', String(ora));
  }
  if (ora - Number(p.getProperty('ultimoRiepilogo') || 0) >= ORE_RIEPILOGO_BIGLIETTERIA * 3600e3 - 5 * 60e3) {
    riepilogoBiglietteria(); p.setProperty('ultimoRiepilogo', String(ora));
  }
}

/* Ordini nuovi (mai notificati): una mail sola, solo se ce ne sono */
function notificaOrdini() {
  const sh = foglio(), righe = dati(sh);
  const nuovi = righe.filter(r => !r.v[C['Notificato il'] - 1] && r.v[C['Stato'] - 1] !== 'Annullato');
  if (!nuovi.length) return;

  const tot = nuovi.reduce((a, r) => a + Number(r.v[C['Totale €'] - 1] || 0), 0);
  const copie = nuovi.reduce((a, r) => a + Number(r.v[C['Copie'] - 1] || 0), 0);
  const html = `<p><b>${nuovi.length} ${nuovi.length === 1 ? 'nuovo ordine' : 'nuovi ordini'}</b> · ${copie} libri · ${euro(tot)}</p>` +
    tabella(['Codice', 'Nome', 'Consegna', 'Libri', 'Totale €', 'Note'], nuovi) +
    `<p><a href="${SpreadsheetApp.getActive().getUrl()}">Apri il foglio degli ordini</a></p>`;
  MailApp.sendEmail({ to: DESTINATARIO, subject: `Libri di Natale · ${nuovi.length} ${nuovi.length === 1 ? 'nuovo ordine' : 'nuovi ordini'}`, htmlBody: html });

  const adesso = new Date();
  nuovi.forEach(r => sh.getRange(r.riga, C['Notificato il']).setValue(adesso));
}

/* Riepilogo per la biglietteria: parte solo se sono arrivati nuovi ritiri dall'ultimo riepilogo */
function riepilogoBiglietteria() {
  const sh = foglio(), righe = dati(sh);
  const daRitirare = righe.filter(r => r.v[C['Consegna'] - 1] === 'Ritiro a Bergamo' && r.v[C['Stato'] - 1] === 'Da ritirare');
  const nuovi = daRitirare.filter(r => !r.v[C['In riepilogo il'] - 1]);
  if (!nuovi.length) return;

  daRitirare.sort((a, b) => String(a.v[C['Nome'] - 1]).localeCompare(String(b.v[C['Nome'] - 1]), 'it'));
  const copie = daRitirare.reduce((a, r) => a + Number(r.v[C['Copie'] - 1] || 0), 0);
  const html = `<p>Riepilogo dei libri da consegnare alla <b>biglietteria del Monastero del Carmine</b>.</p>` +
    `<p><b>${daRitirare.length} ritiri da fare</b> · ${copie} libri · di cui <b>${nuovi.length} nuovi</b> dall'ultimo riepilogo (segnati con ★).</p>` +
    `<p>Il cliente mostra il <b>codice</b> o il QR: consegnare i libri e segnare «Ritirato» nella colonna Stato del foglio.</p>` +
    tabella(['Codice', 'Nome', 'Libri', 'Note'], daRitirare, r => !r.v[C['In riepilogo il'] - 1] ? '★ ' : '');
  MailApp.sendEmail({ to: DESTINATARIO, subject: `Biglietteria · riepilogo ritiri libri (${daRitirare.length} da consegnare)`, htmlBody: html });

  const adesso = new Date();
  nuovi.forEach(r => sh.getRange(r.riga, C['In riepilogo il']).setValue(adesso));
}

/* ---------- utilità ---------- */
function foglio() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(FOGLIO);
  if (!sh) {
    sh = ss.insertSheet(FOGLIO);
    sh.getRange(1, 1, 1, COLONNE.length).setValues([COLONNE]).setFontWeight('bold').setBackground('#f3eadb');
    sh.setFrozenRows(1);
    sh.getRange(2, C['Stato'], 1000, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATI, true).build());
    sh.getRange(2, C['Data'], 1000, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange(2, C['Totale €'], 1000, 1).setNumberFormat('€ #,##0.00');
    sh.setColumnWidth(C['Libri'], 260);
  }
  return sh;
}
function dati(sh) {
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, COLONNE.length).getValues().map((v, i) => ({ riga: i + 2, v }));
}
function tabella(cols, righe, prefisso) {
  const th = cols.map(c => `<th style="text-align:left;padding:6px 10px;border-bottom:2px solid #1e3329">${c}</th>`).join('');
  const tr = righe.map(r => '<tr>' + cols.map((c, j) => {
    let x = r.v[C[c] - 1];
    if (c === 'Totale €') x = euro(x);
    x = esc(x).replace(/\n/g, '<br>');
    if (j === 0 && prefisso) x = prefisso(r) + x;
    return `<td style="padding:6px 10px;border-bottom:1px solid #e6d9c3;vertical-align:top">${x}</td>`;
  }).join('') + '</tr>').join('');
  return `<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px">${'<tr>' + th + '</tr>' + tr}</table>`;
}
function euro(n) { return '€ ' + Number(n || 0).toFixed(2).replace('.', ','); }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

/* ---------- installazione: eseguire UNA volta da Apps Script ---------- */
function prepara() {
  foglio();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'controlloOrario') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('controlloOrario').timeBased().everyHours(1).create();
  const p = PropertiesService.getScriptProperties();
  p.setProperty('ultimaNotifica', String(Date.now()));
  p.setProperty('ultimoRiepilogo', String(Date.now()));
}

/*
 * ISTRUZIONI (dal computer di Casa degli Alfieri, con l'account Google che deve mandare le mail)
 *  1. Crea un nuovo Foglio Google e chiamalo «Ordini libri di Natale».
 *  2. Menu Estensioni → Apps Script. Cancella il codice di esempio e incolla tutto questo file.
 *  3. Cambia CHIAVE (riga 16) con una parola segreta a tua scelta. Salva.
 *  4. In alto scegli la funzione «prepara» e premi Esegui. Autorizza l'accesso quando Google lo chiede.
 *     Si crea il foglio «Ordini» e il controllo automatico ogni ora.
 *  5. Esegui il deployment → Nuovo deployment → tipo «App web».
 *     Esegui come: Me. Chi ha accesso: Chiunque. Premi Esegui il deployment e copia l'URL che finisce con /exec.
 *  6. Nel sito, in NEGOZIO: incolla quell'URL in ordiniURL e la parola segreta in ordiniChiave.
 *
 * Note: il controllo gira ogni ora ma manda le mail solo ogni 3 ore (ordini) e ogni 10 ore (biglietteria),
 * e solo se c'è qualcosa di nuovo. Gli ordini annullati vanno segnati «Annullato» nella colonna Stato.
 */
