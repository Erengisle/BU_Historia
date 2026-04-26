// ========================
// Omprov Historia – Slutversion
// ========================

const SHEET_NAME = 'Omprov Historia';
const TEACHER_EMAIL = 'hakan.hildingsson@edu.huddinge.se';
const CALENDAR_ID = 'primary';

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
