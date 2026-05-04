// ========================
// Omprov Historia – Slutversion
// ========================

const SHEET_NAME     = 'Omprov Historia';
const RESULTAT_SHEET = 'Resultat';
const TEACHER_EMAIL  = 'hakan.hildingsson@edu.huddinge.se';
const CALENDAR_ID    = 'primary';

// Externt blad med namn + e-post
const EXTERN_SHEET_ID   = '1HkzR5BA3uIcG5ZtGgmOO3NHqq5k7pw8REdCXJg9mGHc';
const EXTERN_SHEET_NAMN = 'SPRINT-Hi: Namn och e-postadresser';

// Kolumnpositioner i "Resultat"-sheetet (1-baserat)
// A=1 Namn | B=2 (reserv) | C–N=3–14 Betyg+Kommentar ×6 | O=15 E-post | P=16 Intern | Q=17 Token
const EPOST_KOL      = 15; // Kolumn O
const INTERN_KOM_KOL = 16; // Kolumn P
const TOKEN_KOL      = 17; // Kolumn Q

// URL till detta Apps Script-projekt (omprovformulär + resultatsidor).
// Sätts automatiskt när scriptet är driftsatt som webbapp.
const OMPROV_URL = (function() {
  try { return ScriptApp.getService().getUrl() || ''; } catch(e) { return ''; }
})();

const UPPGIFTER = [
  'Första världskriget',
  'Mellankrigstiden',
  'Andra världskriget',
  'Europa efter världskrigen',
  'Sverige efter världskrigen',
  'Världen efter världskrigen',
];

// ---------- MENY & SIDOPANEL ----------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Historia')
    .addItem('Öppna betygspanel (sidebar)',         'openSidebar')
    .addItem('Öppna betygspanel i webbläsaren',     'visaLararUrl')
    .addItem('Förhandsgranska mejl & resultatsida', 'oppnaForhandsgranskning')
    .addSeparator()
    .addItem('Importera elever från klasslista',    'importeraElever')
    .addItem('Uppdatera e-postadresser',            'uppdateraEpostadresser')
    .addSeparator()
    .addItem('Skicka resultatmejl till alla',       'skickaResultatmail')
    .addSeparator()
    .addItem('Skapa QR-kod (omprov)',               'skapaQRkod')
    .addToUi();
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Betygspanel');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Visar lärarens direktlänk till betygspanelen som webbsida
function visaLararUrl() {
  if (!OMPROV_URL) {
    SpreadsheetApp.getUi().alert(
      'Scriptet är inte driftsatt som webbapp ännu.\n\n' +
      'Gå till Driftsätt → Hantera driftsättningar → skapa en webbapp.'
    );
    return;
  }
  var lararUrl = OMPROV_URL + '?view=larare';
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:24px;">' +
    '<p style="font-size:13px;color:#556070;margin-bottom:12px;">Bokmärk den här länken för enkel åtkomst till betygspanelen:</p>' +
    '<a href="' + lararUrl + '" target="_blank" ' +
    'style="display:block;padding:12px 16px;background:#0F1B2D;color:#F7F4EE;' +
    'text-decoration:none;border-radius:8px;font-size:13px;word-break:break-all;">' +
    lararUrl + '</a>' +
    '<p style="font-size:12px;color:#96A3B0;margin-top:12px;">Länken kräver att du är inloggad med ditt Google-konto.</p>' +
    '</div>'
  ).setWidth(540).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, 'Betygspanel – lärarvy');
}

function oppnaForhandsgranskning() {
  var html = HtmlService.createHtmlOutputFromFile('preview')
    .setWidth(980)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Förhandsgranskning – Historia');
}

// Anropas från preview.html när det körs inuti Sheets
function getWebAppUrl() {
  try { return ScriptApp.getService().getUrl() || ''; } catch(e) { return ''; }
}

// ---------- IMPORTERA ELEVER ----------
// Lägger till nya elever från det externa bladet. Befintliga rörs ej.
// Genererar token för alla som saknar en.
function importeraElever() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RESULTAT_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(RESULTAT_SHEET);
    var rubrik = ['Namn', ''];
    UPPGIFTER.forEach(function(u) {
      rubrik.push(u + ' – Betyg');
      rubrik.push(u + ' – Kommentar');
    });
    rubrik.push('E-post');        // O
    rubrik.push('Intern kommentar'); // P
    rubrik.push('Token');            // Q
    sheet.appendRow(rubrik);
    sheet.setFrozenRows(1);
  } else {
    // Säkerställ rubrikerna för O, P, Q
    sheet.getRange(1, EPOST_KOL).setValue('E-post');
    sheet.getRange(1, INTERN_KOM_KOL).setValue('Intern kommentar');
    sheet.getRange(1, TOKEN_KOL).setValue('Token');

    // Generera tokens för befintliga elever som saknar en
    var befintlig = sheet.getDataRange().getValues();
    for (var r = 1; r < befintlig.length; r++) {
      if (befintlig[r][0] && !befintlig[r][TOKEN_KOL - 1]) {
        sheet.getRange(r + 1, TOKEN_KOL).setValue(genereraToken());
      }
    }
  }

  var externSheet;
  try {
    externSheet = SpreadsheetApp.openById(EXTERN_SHEET_ID)
      .getSheetByName(EXTERN_SHEET_NAMN);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Kunde inte öppna externt blad:\n' + err.message);
    return;
  }
  if (!externSheet) {
    SpreadsheetApp.getUi().alert('Fliken "' + EXTERN_SHEET_NAMN + '" hittades inte.');
    return;
  }

  var externData = externSheet.getDataRange().getValues();

  // Karta över befintliga namn
  var befintligaNamn = {};
  var nuData = sheet.getDataRange().getValues();
  nuData.slice(1).forEach(function(row, idx) {
    if (row[0]) befintligaNamn[String(row[0]).trim().toLowerCase()] = idx + 2;
  });

  var tillagda    = 0;
  var uppdaterade = 0;

  externData.slice(1).forEach(function(row) {
    var namn  = String(row[0] || '').trim();
    var epost = String(row[1] || '').trim();
    if (!namn) return;

    var nyckel = namn.toLowerCase();

    if (befintligaNamn[nyckel]) {
      // Elev finns — fyll i e-post i kolumn O om den saknas
      var radNr   = befintligaNamn[nyckel];
      var aktuell = sheet.getRange(radNr, EPOST_KOL).getValue();
      if (!aktuell && epost) {
        sheet.getRange(radNr, EPOST_KOL).setValue(epost);
        uppdaterade++;
      }
      return;
    }

    // Ny elev — bygg rad med rätt kolumnordning
    // A=namn, B=tom, C–N=betyg/kommentar×6, O=epost, P=internKom, Q=token
    var radData = [namn, ''];
    for (var j = 0; j < UPPGIFTER.length; j++) {
      radData.push('–');  // betyg
      radData.push('');   // kommentar
    }
    radData.push(epost);           // O
    radData.push('');              // P intern kommentar
    radData.push(genereraToken()); // Q token
    sheet.appendRow(radData);
    befintligaNamn[nyckel] = sheet.getLastRow();
    tillagda++;
  });

  var msg = tillagda + ' ny' + (tillagda !== 1 ? 'a' : '') +
    ' elev' + (tillagda !== 1 ? 'er' : '') + ' importerades.';
  if (uppdaterade > 0)
    msg += '\n' + uppdaterade + ' elev' + (uppdaterade !== 1 ? 'er' : '') +
      ' fick e-postadress inlagd (kolumn O).';
  SpreadsheetApp.getUi().alert(msg);
}

// ---------- UPPDATERA E-POSTADRESSER ----------
// Fyller i saknade e-postadresser i kolumn O från det externa bladet.
function uppdateraEpostadresser() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTAT_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet saknas: ' + RESULTAT_SHEET);
    return;
  }

  var externSheet;
  try {
    externSheet = SpreadsheetApp.openById(EXTERN_SHEET_ID)
      .getSheetByName(EXTERN_SHEET_NAMN);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Kunde inte öppna externt blad:\n' + err.message);
    return;
  }

  // Bygg karta extern: namn (lowercase) → e-post
  var externMap = {};
  externSheet.getDataRange().getValues().slice(1).forEach(function(row) {
    var n = String(row[0] || '').trim();
    if (n) externMap[n.toLowerCase()] = String(row[1] || '').trim();
  });

  var data        = sheet.getDataRange().getValues();
  var uppdaterade = 0;
  var saknas      = [];

  for (var i = 1; i < data.length; i++) {
    var namn = String(data[i][0] || '').trim();
    if (!namn) continue;
    if (data[i][EPOST_KOL - 1]) continue; // e-post finns redan

    var epost = externMap[namn.toLowerCase()];
    if (epost) {
      sheet.getRange(i + 1, EPOST_KOL).setValue(epost);
      uppdaterade++;
    } else {
      saknas.push(namn);
    }
  }

  var msg = uppdaterade + ' elev' + (uppdaterade !== 1 ? 'er' : '') +
    ' fick e-postadress inlagd i kolumn O.';
  if (saknas.length)
    msg += '\n\nHittades inte i klasslistan (kontrollera stavning):\n' + saknas.join('\n');
  SpreadsheetApp.getUi().alert(msg);
}

function genereraToken() {
  return Utilities.getUuid();
}

// ---------- BACKEND: HÄMTA ELEVDATA ----------
function hamtaElevdata() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RESULTAT_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(RESULTAT_SHEET);
      var headers = ['Namn', ''];
      UPPGIFTER.forEach(function(u) {
        headers.push(u + ' – Betyg');
        headers.push(u + ' – Kommentar');
      });
      headers.push('E-post');           // O
      headers.push('Intern kommentar'); // P
      headers.push('Token');            // Q
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }

    var data   = sheet.getDataRange().getValues();
    var elever = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;

      var betyg     = [];
      var kommentar = [];
      for (var j = 0; j < UPPGIFTER.length; j++) {
        betyg.push(row[2 + j * 2] || '–');
        kommentar.push(row[3 + j * 2] || '');
      }

      elever.push({
        radIndex:        i,
        namn:            row[0],
        epost:           row[EPOST_KOL - 1]      || '',
        betyg:           betyg,
        kommentar:       kommentar,
        internKommentar: row[INTERN_KOM_KOL - 1] || '',
      });
    }

    return { elever: elever, uppgifter: UPPGIFTER };

  } catch (err) {
    return { fel: err.message };
  }
}

// ---------- BACKEND: SPARA ALLA ELEVER ----------
// Skriver betyg (C–N) och intern kommentar (P). Rör ej e-post (O) eller token (Q).
function sparaAllaElever(elever) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RESULTAT_SHEET);
  if (!sheet) throw new Error('Sheet saknas: ' + RESULTAT_SHEET);

  elever.forEach(function(elev) {
    var rowNum = elev.radIndex + 1;

    // Betyg och kommentarer: kolumn C–N (startkolumn 3, 12 värden)
    var gradeData = [];
    for (var j = 0; j < UPPGIFTER.length; j++) {
      gradeData.push(elev.betyg[j]     || '–');
      gradeData.push(elev.kommentar[j] || '');
    }
    sheet.getRange(rowNum, 3, 1, gradeData.length).setValues([gradeData]);

    // Intern kommentar: kolumn P
    sheet.getRange(rowNum, INTERN_KOM_KOL).setValue(elev.internKommentar || '');
  });
}

// ---------- SKICKA MEJL TILL EN ELEV ----------
function skickaMailTillElev(radIndex) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RESULTAT_SHEET);
  if (!sheet) throw new Error('Sheet saknas: ' + RESULTAT_SHEET);

  var data  = sheet.getDataRange().getValues();
  var row   = data[radIndex + 1];
  var namn  = row[0];
  var epost = row[EPOST_KOL - 1];
  var token = row[TOKEN_KOL - 1];

  if (!epost) throw new Error(
    'Eleven saknar e-postadress (kolumn O).\n' +
    'Kör Historia → Uppdatera e-postadresser.'
  );

  if (!token) {
    token = genereraToken();
    sheet.getRange(radIndex + 2, TOKEN_KOL).setValue(token);
  }

  var betyg     = [];
  var kommentar = [];
  for (var j = 0; j < UPPGIFTER.length; j++) {
    betyg.push(row[2 + j * 2] || '–');
    kommentar.push(row[3 + j * 2] || '');
  }

  var resultUrl = (OMPROV_URL && token) ? OMPROV_URL + '?t=' + token : null;
  var html = buildHtmlEmail(namn, betyg, kommentar, resultUrl);
  MailApp.sendEmail({ to: epost, subject: 'Dina resultat i Historia', htmlBody: html });

  return namn;
}

// ---------- SKICKA RESULTATMEJL TILL ALLA ----------
function skickaResultatmail() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RESULTAT_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet saknas: ' + RESULTAT_SHEET);
    return;
  }

  var data     = sheet.getDataRange().getValues();
  var skickade = 0;
  var hoppade  = 0;
  var fel      = [];

  for (var i = 1; i < data.length; i++) {
    var row   = data[i];
    var namn  = row[0];
    var epost = row[EPOST_KOL - 1];
    var token = row[TOKEN_KOL - 1];
    if (!namn) continue;
    if (!epost) { hoppade++; continue; }

    if (!token) {
      token = genereraToken();
      sheet.getRange(i + 1, TOKEN_KOL).setValue(token);
    }

    var betyg     = [];
    var kommentar = [];
    for (var j = 0; j < UPPGIFTER.length; j++) {
      betyg.push(row[2 + j * 2] || '–');
      kommentar.push(row[3 + j * 2] || '');
    }

    try {
      var resultUrl = (OMPROV_URL && token) ? OMPROV_URL + '?t=' + token : null;
      var html = buildHtmlEmail(namn, betyg, kommentar, resultUrl);
      MailApp.sendEmail({ to: epost, subject: 'Dina resultat i Historia', htmlBody: html });
      skickade++;
    } catch (err) {
      fel.push(namn + ': ' + err.message);
    }
  }

  var msg = 'Skickade mejl till ' + skickade + ' elev' + (skickade !== 1 ? 'er' : '') + '.';
  if (hoppade > 0)
    msg += '\n\n' + hoppade + ' elev' + (hoppade !== 1 ? 'er' : '') +
      ' hoppades över (saknar e-post i kolumn O).\n' +
      'Kör Historia → Uppdatera e-postadresser för att fylla i dem.';
  if (fel.length > 0) msg += '\n\nFel:\n' + fel.join('\n');

  SpreadsheetApp.getUi().alert(msg);
}

// ---------- VISA FORMULÄR, RESULTATSIDA ELLER LÄRARVY ----------
function doGet(e) {
  var params = (e && e.parameter) || {};

  if (params.t)              return serveResultatSida(params.t);
  if (params.view === 'larare') return serveLararsida();

  return HtmlService.createHtmlOutputFromFile('form.html')
    .setTitle('Anmälan till omprov – Historia');
}

// Elevens personliga resultatsida (via token)
function serveResultatSida(token) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTAT_SHEET);
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[TOKEN_KOL - 1]) === token) {
        var betyg = [], kommentar = [];
        for (var j = 0; j < UPPGIFTER.length; j++) {
          betyg.push(row[2 + j * 2] || '–');
          kommentar.push(row[3 + j * 2] || '');
        }
        return HtmlService.createHtmlOutput(buildResultatSida(row[0], betyg, kommentar))
          .setTitle('Dina resultat – Historia')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    }
  }
  return HtmlService.createHtmlOutput(
    '<body style="margin:0;font-family:sans-serif;padding:48px 24px;color:#556070;">' +
    '<p style="font-size:15px;">Länken är ogiltig eller har gått ut.<br>Kontakta din lärare.</p></body>'
  ).setTitle('Ogiltig länk');
}

// Lärarens betygspanel som helsida (samma sidebar.html, men i webbläsaren)
function serveLararsida() {
  return HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Betygspanel – Historia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- TA EMOT OMPROVS-ANMÄLAN ----------

// Anropas via google.script.run från formuläret (undviker CORS-problem med fetch/POST)
function registreraOmprov(email, exam, date) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet saknas: ' + SHEET_NAME);

  const [y, m, d] = date.split('-').map(Number);
  const dateObj   = new Date(y, m - 1, d);
  if (isNaN(dateObj.getTime())) throw new Error('Ogiltigt datum: ' + date);

  const week = getISOWeekNumber(dateObj);

  sheet.appendRow([new Date(), email, date, exam, 'v' + week]);

  const subject = 'Bekräftelse – Omprov i Historia';
  const body    =
    'Hej!\n\nDu är nu anmäld till omprov i Historia.\n\n' +
    'Prov: '  + exam + '\n' +
    'Datum: ' + date + '\n' +
    'Vecka: v' + week + '\n' +
    'Tid: 14:45–16:15\n\nVälkommen!\n/Håkan Hildingsson';

  MailApp.sendEmail(email,        subject, body);
  MailApp.sendEmail(TEACHER_EMAIL, subject, body);

  const cal   = CalendarApp.getCalendarById(CALENDAR_ID);
  const start = new Date(date + 'T14:45:00');
  const end   = new Date(date + 'T16:15:00');
  cal.createEvent('Omprov Historia – ' + exam, start, end,
    { guests: email + ',' + TEACHER_EMAIL, sendInvites: true });
}

// doPost behålls som reserv (t.ex. direktanrop utanför webbläsaren)
function doPost(e) {
  try {
    if (!e || !e.postData) throw new Error('Ingen postData mottagen');
    const data = JSON.parse(e.postData.contents);
    registreraOmprov(data.email, data.exam, data.date);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('FEL: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------- ISO-VECKONUMMER ----------
function getISOWeekNumber(date) {
  const d      = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

// ---------- SKAPA QR-KOD ----------
function skapaQRkod() {
  const ss  = SpreadsheetApp.getActive();
  const url = ScriptApp.getService().getUrl();
  const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=' +
    encodeURIComponent(url);
  const blob = UrlFetchApp.fetch(qrUrl).getBlob().setName('qr.png');

  let sheet = ss.getSheetByName('QR-kod');
  if (!sheet) sheet = ss.insertSheet('QR-kod');
  sheet.clear();
  sheet.insertImage(blob, 1, 1);
  sheet.getRange('A20').setValue('Skanna QR-koden för att anmäla dig');
  sheet.getRange('A21').setValue(url);
}
