// ========================
// Omprov Historia – Slutversion
// ========================

const SHEET_NAME    = 'Omprov Historia';
const RESULTAT_SHEET = 'Resultat';
const TEACHER_EMAIL = 'hakan.hildingsson@edu.huddinge.se';
const CALENDAR_ID   = 'primary';

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
// Tar emot state.elever från sidopanelen och skriver tillbaka till Sheetet
function sparaAllaElever(elever) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RESULTAT_SHEET);
  if (!sheet) throw new Error('Sheet saknas: ' + RESULTAT_SHEET);

  elever.forEach(function(elev) {
    var rowNum  = elev.radIndex + 1;   // getRange är 1-baserat
    var rowData = [elev.namn, elev.epost || ''];
    for (var j = 0; j < UPPGIFTER.length; j++) {
      rowData.push(elev.betyg[j]     || '–');
      rowData.push(elev.kommentar[j] || '');
    }
    rowData.push(elev.internKommentar || '');
    sheet.getRange(rowNum, 1, 1, rowData.length).setValues([rowData]);
  });
}

// ---------- SKICKA RESULTATMEJL ----------
// Skickar HTML-mejl till alla elever i Resultat-sheetet som har en e-postadress
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
    var row  = data[i];
    var namn = row[0];
    var epost = row[1];
    if (!namn) continue;
    if (!epost) { hoppade++; continue; }

    var betyg     = [];
    var kommentar = [];
    for (var j = 0; j < UPPGIFTER.length; j++) {
      betyg.push(row[2 + j * 2] || '–');
      kommentar.push(row[3 + j * 2] || '');
    }

    try {
      var html = buildHtmlEmail(namn, betyg, kommentar);
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

  var sammanfattning = 'Klart! Skickade mejl till ' + skickade + ' elev' + (skickade !== 1 ? 'er' : '') + '.';
  if (hoppade > 0) sammanfattning += '\n(' + hoppade + ' elev' + (hoppade !== 1 ? 'er' : '') + ' saknade e-postadress och hoppades över.)';
  if (fel.length > 0) sammanfattning += '\n\nFel:\n' + fel.join('\n');

  SpreadsheetApp.getUi().alert(sammanfattning);
}

// ---------- VISA FORMULÄR ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('form.html')
    .setTitle('Anmälan till omprov – Historia');
}

// ---------- TA EMOT DATA ----------
function doPost(e) {
  try {
    if (!e || !e.postData) throw new Error("Ingen postData mottagen");

    Logger.log("RAW: " + e.postData.contents);
    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("Sheet saknas: " + SHEET_NAME);

    // ---- Datum (felsäkert) ----
    const [y, m, d] = data.date.split("-").map(Number);
    const dateObj = new Date(y, m - 1, d);

    if (isNaN(dateObj)) {
      throw new Error("Ogiltigt datum: " + data.date);
    }

    // ---- ISO-vecka ----
    const week = getISOWeekNumber(dateObj);

    // ---- Spara i arket ----
    sheet.appendRow([
      new Date(),        // A – tidsstämpel
      data.email,        // B – e-post
      data.date,         // C – datum (TEXT yyyy-mm-dd)
      data.exam,         // D – prov
      "v" + week         // E – veckonummer
    ]);

    // ---- E-post ----
    const subject = "Bekräftelse – Omprov i Historia";
    const body = `
Hej!

Du är nu anmäld till omprov i Historia.

Prov: ${data.exam}
Datum: ${data.date}
Vecka: v${week}
Tid: 14:45–16:15

Välkommen!
/Håkan Hildingsson
`;

    MailApp.sendEmail(data.email, subject, body);
    MailApp.sendEmail(TEACHER_EMAIL, subject, body);

    // ---- Kalender (valfritt men stabilt) ----
    const cal = CalendarApp.getCalendarById(CALENDAR_ID);
    const start = new Date(`${data.date}T14:45:00`);
    const end   = new Date(`${data.date}T16:15:00`);

    cal.createEvent(
      `Omprov Historia – ${data.exam}`,
      start,
      end,
      { guests: `${data.email},${TEACHER_EMAIL}`, sendInvites: true }
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
  const ss = SpreadsheetApp.getActive();
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
