// ============================================================
// skapaEllerUppdateraElevdokument — UPPDATERAD
// Ändringar vs original:
//   • Räknaren (G8:G9 och I8) skrivs automatiskt — du behöver
//     inte fylla i kolumn Q/R i moderdokumentet längre
//   • HTML-mejl skickas till eleven vid skapande av nytt dokument
//   • Varningsflagga skickas med om de två första proven är F
// ============================================================

// ── Kursens provnamn (ändra om kursen förändras) ────────────
const OMRADEN = [
  'Första världskriget',
  'Mellankrigstiden',
  'Andra världskriget',
  'Europa efter världskrigen',
  'Sverige efter världskrigen',
  'Världen efter världskrigen',
];

// ── URL till omprovs-formuläret (fyll i din Apps Script-URL) ─
const OMPROV_URL = 'KLISTRA_IN_DIN_APPS_SCRIPT_URL_HÄR';

// ── Färger ──────────────────────────────────────────────────
const COLOR_RED    = '#B52020';
const COLOR_RED_BG = '#FDEAEA';
const COLOR_GREEN    = '#1A7A4A';
const COLOR_GREEN_BG = '#E8F5EE';
const COLOR_YELLOW    = '#A07800';
const COLOR_YELLOW_BG = '#FDF8D0';

// ============================================================
// updateSheet — skriver betyg, kommentarer och räknare
// ============================================================
function updateSheet(sheetId, studentName, results, comments) {
  if (!sheetId) {
    Logger.log(`Ogiltigt kalkylblads-ID för student: ${studentName}.`);
    return;
  }

  try {
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var sheet = spreadsheet.getActiveSheet();

    // ── Betyg → rad 3 (B–G) ─────────────────────────────────
    results.forEach((result, i) => {
      sheet.getRange(3, i + 2).setValue(result || '-');
    });

    // ── Kommentarer → rad 4 (B–G) ───────────────────────────
    comments.forEach((comment, i) => {
      if (comment && comment.trim() !== '') {
        sheet.getRange(4, i + 2).setValue(comment);
      }
    });

    // ── Räkna godkända (E, C eller A) ───────────────────────
    var count = results.filter(val => ['E', 'C', 'A'].includes(val)).length;

    // ── G7 (merged G7:I7): visa räknaren som siffra ─────────
    sheet.getRange(7, 7).setValue(count);

    // ── Trafikljusfärg på G7 och I7 ─────────────────────────
    var cellG7 = sheet.getRange(7, 7);
    var cellI7 = sheet.getRange(7, 9);
    if (count < 3) {
      cellG7.setBackground(COLOR_RED_BG);
      cellI7.setBackground(COLOR_RED_BG);
    } else if (count === 3) {
      cellG7.setBackground(COLOR_YELLOW_BG);
      cellI7.setBackground(COLOR_YELLOW_BG);
    } else {
      cellG7.setBackground(COLOR_GREEN_BG);
      cellI7.setBackground(COLOR_GREEN_BG);
    }

    // ── C: Skriv bråket automatiskt i G8:G9 och I8 ──────────
    // G8 (sammanslaget med G9) visar täljaren + snedstreck
    sheet.getRange(8, 7).setValue(count + '/');
    // I8 visar nämnaren
    sheet.getRange(8, 9).setValue(6);

    // Färglägg bråkcellerna på samma sätt som G7
    var fracG = sheet.getRange(8, 7, 2, 1); // G8:G9
    var fracI8 = sheet.getRange(8, 9);
    var fracI9 = sheet.getRange(9, 9);
    var fracBg = count < 3 ? COLOR_RED_BG : (count === 3 ? COLOR_YELLOW_BG : COLOR_GREEN_BG);
    fracG.setBackground(fracBg);
    fracI8.setBackground(fracBg);
    fracI9.setBackground(fracBg);

    Logger.log(`Uppdaterat kalkylblad för ${studentName}: ${count}/6 godkända.`);
  } catch (error) {
    Logger.log(`Fel vid uppdatering för ${studentName}: ${error.message}`);
  }
}

// ============================================================
// createOrUpdateStudentDocuments — huvudfunktion (oförändrad logik,
// men utan additionalData och med HTML-mejl vid nyskapande)
// ============================================================
function createOrUpdateStudentDocuments() {

  // Kopiera namn från grupplista
  var sourceNamesSheet = SpreadsheetApp.openById('1HkzR5BA3uIcG5ZtGgmOO3NHqq5k7pw8REdCXJg9mGHc')
    .getSheetByName('SPRINT-Hi: Namn och e-postadresser');
  var names = sourceNamesSheet.getRange(2, 1, sourceNamesSheet.getLastRow() - 1, 1).getValues();

  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Samlade resultat');
  targetSheet.getRange(2, 1, names.length, 1).setValues(names);

  var sheet = targetSheet;
  if (!sheet) throw new Error("Fliken 'Samlade resultat' hittades inte.");

  // E-postadresser
  var emailSheet = SpreadsheetApp.openById('1HkzR5BA3uIcG5ZtGgmOO3NHqq5k7pw8REdCXJg9mGHc')
    .getSheetByName('SPRINT-Hi: Namn och e-postadresser');
  var emailData = emailSheet.getDataRange().getValues();
  var emailMap = {};
  emailData.slice(1).forEach(row => { emailMap[row[0]] = row[1]; });

  var data = sheet.getDataRange().getValues();
  var rows = data.slice(1);

  var templateSheetId = '1ozNLTN_kJPm1_1lV4vwTbiQBFDsIjm5LVaRZGQ0CAuM';
  var folderId = '1KnokYu48l6yHXXi8b5zMgJQTtTPlXxFp';
  var folder = DriveApp.getFolderById(folderId);
  var templateFile = DriveApp.getFileById(templateSheetId);

  rows.forEach(function(row, index) {
    try {
      var studentName = row[0];
      var sheetId = row[15] ? row[15].trim() : '';

      if (!studentName) return;

      var results  = row.slice(1, 7).map(v => v || '-');
      var comments = row.slice(8, 14).map(v => v || '');
      var email    = emailMap[studentName];

      if (!sheetId || sheetId === 'undefined') {
        // ── Skapa nytt dokument (ingen tidigare fil) ───────
        var studentFile = templateFile.makeCopy(`${studentName} - Bedömningar`, folder);
        sheetId = studentFile.getId();
        sheet.getRange(index + 2, 16).setValue(sheetId);
        Logger.log(`Nytt dokument skapat för ${studentName}: ${sheetId}`);
      } else {
        // ── Ta bort gammalt dokument och skapa nytt ────────
        try {
          DriveApp.getFileById(sheetId).setTrashed(true);
          Logger.log(`Gammalt dokument borttaget för ${studentName}: ${sheetId}`);
        } catch(e) {
          Logger.log(`Kunde inte ta bort gammalt dokument för ${studentName}: ${e.message}`);
        }
        var studentFile = templateFile.makeCopy(`${studentName} - Bedömningar`, folder);
        sheetId = studentFile.getId();
        sheet.getRange(index + 2, 16).setValue(sheetId);
        Logger.log(`Nytt dokument skapat för ${studentName}: ${sheetId}`);
      }

      updateSheet(sheetId, studentName, results, comments);

      if (email) {
        studentFile.addEditor(email);
        var htmlBody = buildHtmlEmail(studentName, results, comments);
        GmailApp.sendEmail(
          email,
          `Dina resultat i Historia — ${studentName}`,
          'Se bifogat meddelande (kräver HTML-stöd).',
          { htmlBody: htmlBody }
        );
      }
    } catch (error) {
      Logger.log(`Fel på rad ${index + 2} för '${row[0] || 'Okänd'}': ${error.message}`);
    }
  });
}
