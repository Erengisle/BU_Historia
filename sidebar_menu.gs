// ============================================================
// sidebar_menu.gs
// Lägg in i samma Apps Script-projekt som övriga .gs-filer.
//
// Sheet-struktur ("Samlade resultat"):
//   A  = Elevnamn
//   B  = E-post
//   C  = Betyg uppgift 1
//   D  = Betyg uppgift 2
//   E  = Betyg uppgift 3
//   F  = Kommentar uppgift 1 (skickas till eleven)
//   G  = Kommentar uppgift 2
//   H  = Kommentar uppgift 3
//   I  = Intern anteckning (visas ej för eleven)
//   J  = Google Sheet-ID (elevens dokument)
// ============================================================

const SHEET_SAMLADE = 'Samlade resultat';
const UPPGIFTER_NAMN = [
  'Första världskriget',
  'Mellankrigstiden',
  'Andra världskriget',
];

// ── Meny ────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Historia')
    .addItem('📋 Öppna betygspanel', 'oppnaBetygspanel')
    .addSeparator()
    .addItem('📧 Skicka mejl till alla elever', 'createOrUpdateStudentDocuments')
    .addToUi();
}

// ── Öppna sidopanel ─────────────────────────────────────────
function oppnaBetygspanel() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Betygspanel — Historia')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Hämta alla elever och deras data ────────────────────────
function hamtaElevdata() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SAMLADE);
  if (!sheet) return { fel: 'Fliken "' + SHEET_SAMLADE + '" hittades inte.' };

  var data = sheet.getDataRange().getValues();
  var elever = [];

  data.slice(1).forEach(function(row, i) {
    if (!row[0]) return; // hoppa över tomma rader
    elever.push({
      rad:             i + 2,          // 1-indexerad radnr i Sheet
      namn:            row[0] || '',
      email:           row[1] || '',
      betyg:           [row[2] || '–', row[3] || '–', row[4] || '–'],
      kommentar:       [row[5] || '', row[6] || '', row[7] || ''],
      internKommentar: row[8] || '',
    });
  });

  return { elever: elever, uppgifter: UPPGIFTER_NAMN };
}

// ── Spara en elevs data ──────────────────────────────────────
function sparaElev(rad, betyg, kommentar, internKommentar) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SAMLADE);
  if (!sheet) throw new Error('Fliken "' + SHEET_SAMLADE + '" hittades inte.');

  // Betyg: C–E
  sheet.getRange(rad, 3, 1, 3).setValues([betyg]);
  // Kommentarer: F–H
  sheet.getRange(rad, 6, 1, 3).setValues([kommentar]);
  // Intern anteckning: I
  sheet.getRange(rad, 9).setValue(internKommentar);

  return 'OK';
}

// ── Spara alla elever på en gång ─────────────────────────────
function sparaAllaElever(elever) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SAMLADE);
  if (!sheet) throw new Error('Fliken "' + SHEET_SAMLADE + '" hittades inte.');

  elever.forEach(function(e) {
    sheet.getRange(e.rad, 3, 1, 3).setValues([e.betyg]);
    sheet.getRange(e.rad, 6, 1, 3).setValues([e.kommentar]);
    sheet.getRange(e.rad, 9).setValue(e.internKommentar || '');
  });

  return 'OK';
}
