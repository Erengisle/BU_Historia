// ============================================================
// buildHtmlEmail — genererar mobilanpassat HTML-mejl
// Används av skapaEllerUppdateraElevdokument_v2.gs
//
// Lägg båda filerna i samma Apps Script-projekt.
// ============================================================

function buildHtmlEmail(studentName, results, comments) {

  // ── Hjälpfunktioner ─────────────────────────────────────
  function gradeColor(g) {
    if (g === 'A') return { bg: '#E8F5EE', color: '#1A7A4A' };
    if (g === 'C') return { bg: '#EBF0FA', color: '#2355A0' };
    if (g === 'E') return { bg: '#FDF0E8', color: '#C05A20' };
    return { bg: '#F0EDE6', color: '#96A3B0' };
  }

  function gradeBadge(g, size) {
    size = size || 36;
    var c = gradeColor(g);
    return '<span style="display:inline-flex;align-items:center;justify-content:center;' +
      'width:' + size + 'px;height:' + size + 'px;border-radius:6px;' +
      'background:' + c.bg + ';font-family:Georgia,serif;font-size:' + Math.round(size * 0.54) + 'px;' +
      'font-weight:700;color:' + c.color + ';line-height:1;">' + g + '</span>';
  }

  var count = results.filter(function(r) { return ['E','C','A'].indexOf(r) !== -1; }).length;

  // Trafikljus
  var tcBg    = count >= 4 ? '#E8F5EE' : (count === 3 ? '#FDF8D0' : '#FDEAEA');
  var tcColor = count >= 4 ? '#1A7A4A' : (count === 3 ? '#A07800' : '#B52020');
  var tcLabel = count >= 4 ? 'Du klarar kursen' : (count === 3 ? 'Nästan — ett prov kvar' : 'Fler godkända behövs');

  // Visa varning om de två första proven är F/-
  // Varningsnivå baserat på antal F och om de två första proven misslyckades
  var first2Fail = !['E','C','A'].includes(results[0]) && !['E','C','A'].includes(results[1]);
  var totalFail  = results.filter(function(r) { return !['E','C','A'].includes(r); }).length;
  var warnLevel  = 0;
  if (first2Fail) {
    if (totalFail >= 4)      warnLevel = 3;
    else if (totalFail >= 3) warnLevel = 2;
    else                     warnLevel = 1;
  }

  // ── Varningsbanner ────────────────────────────────────
  var warningHtml = '';
  if (warnLevel > 0) {
    var warnTitle, warnBody, ctaLabel, ctaHref, ctaBg, ctaColor, tipsHtml;
    var tips =
      '<table cellpadding="0" cellspacing="0" style="width:100%;margin:10px 0 14px;">' +
        '<tr><td style="padding:6px 0;font-size:13px;color:#0F1B2D;">📚 <strong>Har du planerat dina studier?</strong> Sprid ut läsningen — plugga inte bara dagen innan.</td></tr>' +
        '<tr><td style="padding:6px 0;font-size:13px;color:#0F1B2D;">💻 <strong>Har du tittat på materialet i Google Classroom?</strong> Presentationer och filmer finns där.</td></tr>' +
        '<tr><td style="padding:6px 0;font-size:13px;color:#0F1B2D;">📖 <strong>Har du gjort dina läxor?</strong> Det gör lektionerna lättare att hänga med på.</td></tr>' +
        '<tr><td style="padding:6px 0;font-size:13px;color:#0F1B2D;">❓ <strong>Har du frågat om du inte förstått?</strong> Det är aldrig fel att höra av sig.</td></tr>' +
      '</table>';

    if (warnLevel === 1) {
      warnTitle = 'Du har inte klarat de två första proven ännu';
      warnBody  = 'Det är helt okej — men det är bra att börja tänka på din studieteknik och eventuellt boka in ett omprov.';
      ctaLabel  = 'Anmäl dig till omprov →';
      ctaHref   = OMPROV_URL;
      ctaBg = '#B52020'; ctaColor = '#fff';
      tipsHtml  = tips;
    } else if (warnLevel === 2) {
      warnTitle = 'Du har inte klarat tre prov — det är dags att agera';
      warnBody  = 'För att klara kursen krävs minst 4 godkända prov. Du behöver höra av dig till mig så vi kan lägga upp en plan tillsammans.';
      ctaLabel  = 'Anmäl dig till omprov →';
      ctaHref   = OMPROV_URL;
      ctaBg = '#B52020'; ctaColor = '#fff';
      tipsHtml  = tips;
    } else {
      warnTitle = 'Du riskerar att inte klara kursen';
      warnBody  = 'Du har inte klarat fyra eller fler prov. Det är nu mycket viktigt att du hör av dig till mig — vi måste prata om vad som krävs för att du ska ha en chans att klara kursen.';
      ctaLabel  = 'Kontakta läraren direkt →';
      ctaHref   = 'mailto:' + TEACHER_EMAIL;
      ctaBg = '#F7F4EE'; ctaColor = '#7A1010';
      tipsHtml  = '';
    }

    warningHtml =
      '<table width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 16px;border-radius:10px;overflow:hidden;border:1px solid #E8AAAA;">' +
        '<tr><td style="background:#FDEAEA;padding:16px;">' +
          '<p style="font-family:Georgia,serif;font-size:16px;font-weight:700;color:#7A1010;margin:0 0 6px;">' + warnTitle + '</p>' +
          '<p style="font-size:13px;color:#8B2020;margin:0;line-height:1.6;">' + warnBody + '</p>' +
          tipsHtml +
          '<a href="' + ctaHref + '" style="display:block;padding:12px 16px;background:' + ctaBg + ';color:' + ctaColor + ';text-decoration:none;border-radius:8px;font-size:14px;font-weight:700;text-align:center;">' + ctaLabel + '</a>' +
        '</td></tr>' +
      '</table>';
  }

  // ── Räknare ───────────────────────────────────────────
  var counterHtml =
    '<table width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 16px;background:' + tcBg + ';border-radius:10px;">' +
      '<tr>' +
        '<td style="padding:14px 18px;">' +
          '<span style="font-family:Georgia,serif;font-size:32px;font-weight:700;color:' + tcColor + ';">' + count + '/6</span>' +
          '<span style="font-size:13px;color:' + tcColor + ';font-weight:600;margin-left:10px;">' + tcLabel + '</span>' +
        '</td>' +
      '</tr>' +
    '</table>';

  // ── Sätt ihop hela mejlet ─────────────────────────────
  return '<!DOCTYPE html>' +
  '<html lang="sv"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
  '<body style="margin:0;padding:0;background:#F0EDE6;font-family:Helvetica Neue,Arial,sans-serif;">' +

    // Wrapper
    '<table width="100%" cellpadding="0" cellspacing="0">' +
    '<tr><td align="center" style="padding:20px 12px 40px;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;">' +

      // Header
      '<tr><td style="background:#0F1B2D;border-radius:12px 12px 0 0;padding:24px 24px 20px;">' +
        '<p style="margin:0 0 6px;font-size:10px;font-weight:700;letter-spacing:0.12em;text-transform:uppercase;color:#556070;">Historia · Grundskola · Vt 2026</p>' +
        '<p style="margin:0;font-family:Georgia,serif;font-size:24px;font-weight:700;color:#F7F4EE;line-height:1.2;">' + studentName + '</p>' +
        '<p style="margin:6px 0 0;font-size:13px;color:#556070;">Dina resultat och kommentarer</p>' +
      '</td></tr>' +

      // Body
      '<tr><td style="background:#F7F4EE;padding:16px;border-radius:0 0 12px 12px;">' +

        warningHtml +
        counterHtml +

        // Prov-tabell
        '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:10px;border:1px solid #DDD8D0;overflow:hidden;margin-bottom:16px;">' +
          provRows +
        '</table>' +

        // Info-not
        '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:10px;border:1px solid #DDD8D0;margin-bottom:8px;">' +
        '<tr><td style="padding:14px 16px;">' +
          '<p style="margin:0;font-family:Georgia,serif;font-style:italic;font-size:13px;color:#556070;line-height:1.75;">' +
            'För godkänt på kursen krävs 4 av 6 godkända prov. Omprov planeras vid behov — alltid onsdagar kl. 14:45–16:15.' +
          '</p>' +
        '</td></tr></table>' +

        // Footer
        '<p style="margin:12px 0 0;font-size:11px;color:#96A3B0;text-align:center;">Historia · Grundskola · Håkan Hildingsson · Vt 2026</p>' +

      '</td></tr>' +
    '</table>' +
    '</td></tr></table>' +

  '</body></html>';
}
