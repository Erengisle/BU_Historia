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

// Token-kolumn = P (kolumn 16, 1-baserat)
const TOKEN_KOL = 16;

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
    .addItem('Öppna betygspanel', 'openSidebar')
    .addSeparator()
    .addItem('Importera elever från klasslista', 'importeraElever')
    .addSeparator()
    .addItem('Skicka resultatmejl till alla', 'skickaResultatmail')
    .addSeparator()
    .addItem('Skapa QR-kod (omprov)', 'skapaQRkod')
    .addToUi();
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Betygspanel');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ---------- IMPORTERA ELEVER ----------
// Läser namn + e-post från det externa bladet och lägger till nya elever
// i "Resultat"-sheetet. Befintliga elever rörs inte (utom att e-post
// fylls i om den saknas). Genererar token för alla som saknar en.
function importeraElever() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RESULTAT_SHEET);

  // Skapa sheet med rätt rubriker om det inte finns
  if (!sheet) {
    sheet = ss.insertSheet(RESULTAT_SHEET);
    var rubrik = ['Namn', 'E-post'];
    UPPGIFTER.forEach(function(u) {
      rubrik.push(u + ' – Betyg');
      rubrik.push(u + ' – Kommentar');
    });
    rubrik.push('Intern kommentar');
    rubrik.push('Token');
    sheet.appendRow(rubrik);
    sheet.setFrozenRows(1);
  } else {
    // Säkerställ att Token-kolumnen finns
    var rubrikRad = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (rubrikRad[TOKEN_KOL - 1] !== 'Token') {
      sheet.getRange(1, TOKEN_KOL).setValue('Token');
    }
    // Generera tokens för befintliga elever som saknar en
    var befintlig = sheet.getDataRange().getValues();
    for (var r = 1; r < befintlig.length; r++) {
      if (befintlig[r][0] && !befintlig[r][TOKEN_KOL - 1]) {
        sheet.getRange(r + 1, TOKEN_KOL).setValue(genereraToken());
      }
    }
  }

  // Öppna externt blad
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

  // Bygg karta över befintliga namn (case-insensitive)
  var befintligaNamn = {};
  var nuData = sheet.getDataRange().getValues();
  nuData.slice(1).forEach(function(row, idx) {
    if (row[0]) befintligaNamn[String(row[0]).trim().toLowerCase()] = idx + 2; // radnr 1-baserat
  });

  var tillagda   = 0;
  var uppdaterade = 0;

  externData.slice(1).forEach(function(row) {
    var namn  = String(row[0] || '').trim();
    var epost = String(row[1] || '').trim();
    if (!namn) return;

    var nyckel = namn.toLowerCase();

    if (befintligaNamn[nyckel]) {
      // Elev finns redan — fyll i e-post om den saknas
      var radNr = befintligaNamn[nyckel];
      var aktuell = sheet.getRange(radNr, 2).getValue();
      if (!aktuell && epost) {
        sheet.getRange(radNr, 2).setValue(epost);
        uppdaterade++;
      }
      return;
    }

    // Ny elev
    var radData = [namn, epost];
    for (var j = 0; j < UPPGIFTER.length; j++) {
      radData.push('–');
      radData.push('');
    }
    radData.push('');             // intern kommentar
    radData.push(genereraToken()); // token
    sheet.appendRow(radData);
    befintligaNamn[nyckel] = sheet.getLastRow();
    tillagda++;
  });

  var msg = 'Klart!\n' + tillagda + ' ny' + (tillagda !== 1 ? 'a' : '') +
    ' elev' + (tillagda !== 1 ? 'er' : '') + ' importerades.';
  if (uppdaterade > 0)
    msg += '\n' + uppdaterade + ' elev' + (uppdaterade !== 1 ? 'er' : '') +
      ' fick e-postadress inlagd.';
  SpreadsheetApp.getUi().alert(msg);
}

function genereraToken() {
  return Utilities.getUuid();
}

// ---------- BACKEND: HÄMTA ELEVDATA ----------
// Returnerar { elever: [...], uppgifter: [...] } eller { fel: "..." }
function hamtaElevdata() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RESULTAT_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(RESULTAT_SHEET);
      var headers = ['Namn', 'E-post'];
      UPPGIFTER.forEach(function(u) {
        headers.push(u + ' – Betyg');
        headers.push(u + ' – Kommentar');
      });
      headers.push('Intern kommentar');
      headers.push('Token');
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
        epost:           row[1] || '',
        betyg:           betyg,
        kommentar:       kommentar,
        internKommentar: row[2 + UPPGIFTER.length * 2] || '',
      });
    }

    return { elever: elever, uppgifter: UPPGIFTER };

  } catch (err) {
    return { fel: err.message };
  }
}

// ---------- BACKEND: SPARA ALLA ELEVER ----------
// Skriver kolumnerna A–O (lämnar token i P orörd)
function sparaAllaElever(elever) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RESULTAT_SHEET);
  if (!sheet) throw new Error('Sheet saknas: ' + RESULTAT_SHEET);

  elever.forEach(function(elev) {
    var rowNum  = elev.radIndex + 1;
    var rowData = [elev.namn, elev.epost || ''];
    for (var j = 0; j < UPPGIFTER.length; j++) {
      rowData.push(elev.betyg[j]     || '–');
      rowData.push(elev.kommentar[j] || '');
    }
    rowData.push(elev.internKommentar || '');
    // Skriver exakt 15 kolumner (A–O) — token i kolumn P rörs ej
    sheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
  });
}

// ---------- SKICKA MEJL TILL EN ELEV ----------
// Anropas från sidopanelen med elevens radIndex (0-baserat, exkl. rubrikrad)
function skickaMailTillElev(radIndex) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RESULTAT_SHEET);
  if (!sheet) throw new Error('Sheet saknas: ' + RESULTAT_SHEET);

  var data  = sheet.getDataRange().getValues();
  var row   = data[radIndex + 1];
  var namn  = row[0];
  var epost = row[1];
  var token = row[TOKEN_KOL - 1];

  if (!epost) throw new Error('Eleven saknar e-postadress.');

  var betyg     = [];
  var kommentar = [];
  for (var j = 0; j < UPPGIFTER.length; j++) {
    betyg.push(row[2 + j * 2] || '–');
    kommentar.push(row[3 + j * 2] || '');
  }

  var resultUrl = token ? ScriptApp.getService().getUrl() + '?t=' + token : null;
  var html = buildHtmlEmail(namn, betyg, kommentar, resultUrl);
  MailApp.sendEmail({
    to:       epost,
    subject:  'Dina resultat i Historia',
    htmlBody: html,
  });

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
    var epost = row[1];
    var token = row[TOKEN_KOL - 1];
    if (!namn) continue;
    if (!epost) { hoppade++; continue; }

    var betyg     = [];
    var kommentar = [];
    for (var j = 0; j < UPPGIFTER.length; j++) {
      betyg.push(row[2 + j * 2] || '–');
      kommentar.push(row[3 + j * 2] || '');
    }

    try {
      var resultUrl = token ? ScriptApp.getService().getUrl() + '?t=' + token : null;
      var html = buildHtmlEmail(namn, betyg, kommentar, resultUrl);
      MailApp.sendEmail({
        to:       epost,
        subject:  'Dina resultat i Historia',
        htmlBody: html,
      });
      skickade++;
    } catch (err) {
      fel.push(namn + ': ' + err.message);
    }
  }

  var sammanfattning = 'Klart! Skickade mejl till ' + skickade + ' elev' +
    (skickade !== 1 ? 'er' : '') + '.';
  if (hoppade > 0)
    sammanfattning += '\n(' + hoppade + ' elev' + (hoppade !== 1 ? 'er' : '') +
      ' saknade e-postadress och hoppades över.)';
  if (fel.length > 0) sammanfattning += '\n\nFel:\n' + fel.join('\n');

  SpreadsheetApp.getUi().alert(sammanfattning);
}

// ---------- VISA FORMULÄR ELLER RESULTATSIDA ----------
function doGet(e) {
  var token = e && e.parameter && e.parameter.t;
  if (token) return serveResultatSida(token);
  return HtmlService.createHtmlOutputFromFile('form.html')
    .setTitle('Anmälan till omprov – Historia');
}

// Letar upp eleven via token och returnerar deras resultatsida
function serveResultatSida(token) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESULTAT_SHEET);
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[TOKEN_KOL - 1]) === token) {
        var namn = row[0];
        var betyg = [], kommentar = [];
        for (var j = 0; j < UPPGIFTER.length; j++) {
          betyg.push(row[2 + j * 2] || '–');
          kommentar.push(row[3 + j * 2] || '');
        }
        return HtmlService.createHtmlOutput(buildResultatSida(namn, betyg, kommentar))
          .setTitle('Dina resultat – Historia')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    }
  }
  return HtmlService.createHtmlOutput(
    '<body style="margin:0;font-family:sans-serif;padding:48px 24px;color:#556070;">' +
    '<p style="font-size:15px;">Länken är ogiltig eller har gått ut.<br>' +
    'Kontakta din lärare.</p></body>'
  ).setTitle('Ogiltig länk');
}

// ---------- TA EMOT OMPROVS-ANMÄLAN ----------
function doPost(e) {
  try {
    if (!e || !e.postData) throw new Error("Ingen postData mottagen");

    Logger.log("RAW: " + e.postData.contents);
    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("Sheet saknas: " + SHEET_NAME);

    const [y, m, d] = data.date.split("-").map(Number);
    const dateObj = new Date(y, m - 1, d);
    if (isNaN(dateObj)) throw new Error("Ogiltigt datum: " + data.date);

    const week = getISOWeekNumber(dateObj);

    sheet.appendRow([
      new Date(),
      data.email,
      data.date,
      data.exam,
      "v" + week,
    ]);

    const subject = "Bekräftelse – Omprov i Historia";
    const body =
      "Hej!\n\nDu är nu anmäld till omprov i Historia.\n\n" +
      "Prov: " + data.exam + "\n" +
      "Datum: " + data.date + "\n" +
      "Vecka: v" + week + "\n" +
      "Tid: 14:45–16:15\n\n" +
      "Välkommen!\n/Håkan Hildingsson";

    MailApp.sendEmail(data.email, subject, body);
    MailApp.sendEmail(TEACHER_EMAIL, subject, body);

    const cal   = CalendarApp.getCalendarById(CALENDAR_ID);
    const start = new Date(data.date + "T14:45:00");
    const end   = new Date(data.date + "T16:15:00");
    cal.createEvent(
      "Omprov Historia – " + data.exam,
      start,
      end,
      { guests: data.email + "," + TEACHER_EMAIL, sendInvites: true }
    );

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("FEL: " + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------- ISO-VECKONUMMER ----------
function getISOWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

// ---------- SKAPA QR-KOD ----------
function skapaQRkod() {
  const ss  = SpreadsheetApp.getActive();
  const url = ScriptApp.getService().getUrl();

  const qrUrl =
    "https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=" +
    encodeURIComponent(url);

  const blob = UrlFetchApp.fetch(qrUrl).getBlob().setName("qr.png");

  let sheet = ss.getSheetByName("QR-kod");
  if (!sheet) sheet = ss.insertSheet("QR-kod");
  sheet.clear();

  sheet.insertImage(blob, 1, 1);
  sheet.getRange("A20").setValue("Skanna QR-koden för att anmäla dig");
  sheet.getRange("A21").setValue(url);
}
