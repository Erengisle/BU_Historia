// ============================================================
// buildHtmlEmail  — mobilanpassat HTML-mejl
// buildResultatSida — elevens personliga resultatsida (webb)
// ============================================================

// ── Delade hjälpfunktioner ──────────────────────────────────

function gradeColor_(g) {
  if (g === 'A') return { bg: '#E8F5EE', color: '#1A7A4A' };
  if (g === 'C') return { bg: '#EBF0FA', color: '#2355A0' };
  if (g === 'E') return { bg: '#FDF0E8', color: '#C05A20' };
  return           { bg: '#F0EDE6', color: '#96A3B0' };
}

function gradeBadge_(g, size) {
  size = size || 36;
  var c = gradeColor_(g);
  return '<span style="display:inline-flex;align-items:center;justify-content:center;' +
    'width:' + size + 'px;height:' + size + 'px;border-radius:6px;' +
    'background:' + c.bg + ';font-family:Georgia,serif;font-size:' + Math.round(size * 0.54) + 'px;' +
    'font-weight:700;color:' + c.color + ';line-height:1;">' + g + '</span>';
}

function warnBlock_(warnLevel, resultUrl) {
  if (warnLevel === 0) return '';

  var tips =
    '<table cellpadding="0" cellspacing="0" style="width:100%;margin:10px 0 14px;">' +
      '<tr><td style="padding:6px 0;font-size:13px;color:#0F1B2D;">📚 <strong>Har du planerat dina studier?</strong> Sprid ut läsningen — plugga inte bara dagen innan.</td></tr>' +
      '<tr><td style="padding:6px 0;font-size:13px;color:#0F1B2D;">💻 <strong>Har du tittat på materialet i Google Classroom?</strong> Presentationer och filmer finns där.</td></tr>' +
      '<tr><td style="padding:6px 0;font-size:13px;color:#0F1B2D;">📖 <strong>Har du gjort dina läxor?</strong> Det gör lektionerna lättare att hänga med på.</td></tr>' +
      '<tr><td style="padding:6px 0;font-size:13px;color:#0F1B2D;">❓ <strong>Har du frågat om du inte förstått?</strong> Det är aldrig fel att höra av sig.</td></tr>' +
    '</table>';

  var warnTitle, warnBody, ctaLabel, ctaHref, ctaBg, ctaColor, tipsHtml;

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

  return '<table width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 16px;border-radius:10px;overflow:hidden;border:1px solid #E8AAAA;">' +
    '<tr><td style="background:#FDEAEA;padding:16px;">' +
      '<p style="font-family:Georgia,serif;font-size:16px;font-weight:700;color:#7A1010;margin:0 0 6px;">' + warnTitle + '</p>' +
      '<p style="font-size:13px;color:#8B2020;margin:0;line-height:1.6;">' + warnBody + '</p>' +
      tipsHtml +
      '<a href="' + ctaHref + '" style="display:block;padding:12px 16px;background:' + ctaBg + ';color:' + ctaColor + ';text-decoration:none;border-radius:8px;font-size:14px;font-weight:700;text-align:center;">' + ctaLabel + '</a>' +
    '</td></tr>' +
  '</table>';
}

function warnLevel_(results) {
  var first2Fail = !['E','C','A'].includes(results[0]) && !['E','C','A'].includes(results[1]);
  if (!first2Fail) return 0;
  var totalFail = results.filter(function(r) { return !['E','C','A'].includes(r); }).length;
  if (totalFail >= 4) return 3;
  if (totalFail >= 3) return 2;
  return 1;
}

// ── buildHtmlEmail ──────────────────────────────────────────
// resultUrl (valfri): länk till elevens personliga resultatsida
function buildHtmlEmail(studentName, results, comments, resultUrl) {

  var count   = results.filter(function(r) { return ['E','C','A'].indexOf(r) !== -1; }).length;
  var tcBg    = count >= 4 ? '#E8F5EE' : (count === 3 ? '#FDF8D0' : '#FDEAEA');
  var tcColor = count >= 4 ? '#1A7A4A' : (count === 3 ? '#A07800' : '#B52020');
  var tcLabel = count >= 4 ? 'Du klarar kursen' : (count === 3 ? 'Nästan — ett prov kvar' : 'Fler godkända behövs');

  var warningHtml = warnBlock_(warnLevel_(results), resultUrl);

  var counterHtml =
    '<table width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 16px;background:' + tcBg + ';border-radius:10px;">' +
      '<tr><td style="padding:14px 18px;">' +
        '<span style="font-family:Georgia,serif;font-size:32px;font-weight:700;color:' + tcColor + ';">' + count + '/6</span>' +
        '<span style="font-size:13px;color:' + tcColor + ';font-weight:600;margin-left:10px;">' + tcLabel + '</span>' +
      '</td></tr>' +
    '</table>';

  var provRows = '';
  results.forEach(function(grade, i) {
    var provNamn = UPPGIFTER[i] || ('Prov ' + (i + 1));
    var comment  = (comments && comments[i]) ? comments[i] : '';
    var isLast   = (i === results.length - 1);
    provRows +=
      '<tr><td style="padding:12px 16px;' + (isLast ? '' : 'border-bottom:1px solid #DDD8D0;') + '">' +
        '<table width="100%" cellpadding="0" cellspacing="0"><tr>' +
          '<td width="40" style="vertical-align:middle;">' + gradeBadge_(grade, 36) + '</td>' +
          '<td style="vertical-align:middle;padding-left:12px;">' +
            '<p style="margin:0;font-size:14px;font-weight:600;color:#0F1B2D;">' + provNamn + '</p>' +
            (comment ? '<p style="margin:4px 0 0;font-size:12px;color:#556070;line-height:1.5;">' + comment + '</p>' : '') +
          '</td>' +
        '</tr></table>' +
      '</td></tr>';
  });

  // Länkknapp till resultatsidan (om URL finns)
  var sidoLankHtml = resultUrl
    ? '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:16px;">' +
        '<tr><td>' +
          '<a href="' + resultUrl + '" style="display:block;padding:13px 16px;background:#0F1B2D;color:#F7F4EE;' +
          'text-decoration:none;border-radius:8px;font-size:14px;font-weight:700;text-align:center;">' +
          'Se din resultatsida →</a>' +
        '</td></tr>' +
      '</table>'
    : '';

  return '<!DOCTYPE html>' +
  '<html lang="sv"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
  '<body style="margin:0;padding:0;background:#F0EDE6;font-family:Helvetica Neue,Arial,sans-serif;">' +
    '<table width="100%" cellpadding="0" cellspacing="0">' +
    '<tr><td align="center" style="padding:20px 12px 40px;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;">' +
      '<tr><td style="background:#0F1B2D;border-radius:12px 12px 0 0;padding:24px 24px 20px;">' +
        '<p style="margin:0 0 6px;font-size:10px;font-weight:700;letter-spacing:0.12em;text-transform:uppercase;color:#556070;">Historia · Grundskola · Vt 2026</p>' +
        '<p style="margin:0;font-family:Georgia,serif;font-size:24px;font-weight:700;color:#F7F4EE;line-height:1.2;">' + studentName + '</p>' +
        '<p style="margin:6px 0 0;font-size:13px;color:#556070;">Dina resultat och kommentarer</p>' +
      '</td></tr>' +
      '<tr><td style="background:#F7F4EE;padding:16px;border-radius:0 0 12px 12px;">' +
        warningHtml +
        counterHtml +
        '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:10px;border:1px solid #DDD8D0;overflow:hidden;margin-bottom:16px;">' +
          provRows +
        '</table>' +
        sidoLankHtml +
        '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:10px;border:1px solid #DDD8D0;margin-bottom:8px;">' +
        '<tr><td style="padding:14px 16px;">' +
          '<p style="margin:0;font-family:Georgia,serif;font-style:italic;font-size:13px;color:#556070;line-height:1.75;">' +
            'För godkänt på kursen krävs 4 av 6 godkända prov. Omprov planeras vid behov — alltid onsdagar kl. 14:45–16:15.' +
          '</p>' +
        '</td></tr></table>' +
        '<p style="margin:12px 0 0;font-size:11px;color:#96A3B0;text-align:center;">Historia · Grundskola · Håkan Hildingsson · Vt 2026</p>' +
      '</td></tr>' +
    '</table>' +
    '</td></tr></table>' +
  '</body></html>';
}

// ── buildResultatSida ───────────────────────────────────────
// Elevens personliga resultatsida — returnerar fullständig HTML-sträng
function buildResultatSida(studentName, results, comments) {

  var count   = results.filter(function(r) { return ['E','C','A'].indexOf(r) !== -1; }).length;
  var tcBg    = count >= 4 ? '#E8F5EE' : (count === 3 ? '#FDF8D0' : '#FDEAEA');
  var tcColor = count >= 4 ? '#1A7A4A' : (count === 3 ? '#A07800' : '#B52020');
  var tcLabel = count >= 4 ? 'Du klarar kursen' : (count === 3 ? 'Nästan — ett prov kvar' : 'Fler godkända behövs');

  var warningHtml = warnBlock_(warnLevel_(results), null);

  // Provresultat-rader
  var provRader = '';
  results.forEach(function(grade, i) {
    var provNamn = UPPGIFTER[i] || ('Prov ' + (i + 1));
    var comment  = (comments && comments[i]) ? comments[i] : '';
    var gc = gradeColor_(grade);
    provRader +=
      '<div class="prov-rad">' +
        '<span class="badge" style="background:' + gc.bg + ';color:' + gc.color + ';">' + grade + '</span>' +
        '<div class="prov-text">' +
          '<div class="prov-namn">' + provNamn + '</div>' +
          (comment ? '<div class="prov-kommentar">' + comment + '</div>' : '') +
        '</div>' +
      '</div>';
  });

  return '<!DOCTYPE html>' +
  '<html lang="sv">' +
  '<head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Dina resultat – Historia</title>' +
    '<link rel="preconnect" href="https://fonts.googleapis.com">' +
    '<link href="https://fonts.googleapis.com/css2?family=Spectral:ital,wght@0,600;1,400&family=Space+Grotesk:wght@400;500;600;700&display=swap" rel="stylesheet">' +
    '<style>' +
      '*{box-sizing:border-box;margin:0;padding:0}' +
      'body{font-family:"Space Grotesk",Helvetica Neue,Arial,sans-serif;background:#F0EDE6;color:#0F1B2D;min-height:100vh;-webkit-font-smoothing:antialiased}' +
      '.header{background:#0F1B2D;padding:28px 20px 22px}' +
      '.header-eyebrow{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#556070;margin-bottom:6px}' +
      '.header-name{font-family:"Spectral",Georgia,serif;font-size:26px;font-weight:600;color:#F7F4EE;line-height:1.2}' +
      '.header-sub{margin-top:6px;font-size:13px;color:#556070}' +
      '.content{padding:16px 14px 48px;max-width:600px;margin:0 auto}' +
      '.trafikljus{background:' + tcBg + ';border-radius:10px;padding:14px 18px;margin-bottom:14px;display:flex;align-items:center;gap:12px}' +
      '.trafikljus-siffra{font-family:"Spectral",Georgia,serif;font-size:34px;font-weight:600;color:' + tcColor + ';line-height:1}' +
      '.trafikljus-label{font-size:13px;font-weight:600;color:' + tcColor + '}' +
      '.kort{background:#fff;border:1px solid #DDD8D0;border-radius:12px;overflow:hidden;margin-bottom:14px}' +
      '.prov-rad{display:flex;align-items:flex-start;gap:12px;padding:13px 16px;border-bottom:1px solid #DDD8D0}' +
      '.prov-rad:last-child{border-bottom:none}' +
      '.badge{display:flex;align-items:center;justify-content:center;width:38px;height:38px;border-radius:7px;font-family:"Spectral",Georgia,serif;font-size:20px;font-weight:600;flex-shrink:0}' +
      '.prov-text{padding-top:2px}' +
      '.prov-namn{font-size:14px;font-weight:600;color:#0F1B2D}' +
      '.prov-kommentar{margin-top:4px;font-size:12px;color:#556070;line-height:1.55}' +
      '.info-kort{background:#fff;border:1px solid #DDD8D0;border-radius:12px;padding:14px 16px;margin-bottom:14px}' +
      '.info-text{font-family:"Spectral",Georgia,serif;font-style:italic;font-size:13px;color:#556070;line-height:1.75}' +
      '.cta-knapp{display:block;width:100%;padding:15px;background:#0F1B2D;color:#F7F4EE;text-decoration:none;border-radius:10px;font-size:15px;font-weight:700;text-align:center;margin-bottom:14px}' +
      '.footer{font-size:11px;color:#96A3B0;text-align:center;padding-top:4px}' +
      // Warning-block återanvänder email-tabellerna (inline styles)
    '</style>' +
  '</head>' +
  '<body>' +
    '<div class="header">' +
      '<div class="header-eyebrow">Historia · Grundskola · Vt 2026</div>' +
      '<div class="header-name">' + studentName + '</div>' +
      '<div class="header-sub">Dina resultat och kommentarer</div>' +
    '</div>' +
    '<div class="content">' +
      // Varningsblock (återanvänder tabell-HTML från buildHtmlEmail)
      (warningHtml ? '<div style="margin-top:0;">' + warningHtml + '</div>' : '') +
      '<div class="trafikljus">' +
        '<div class="trafikljus-siffra">' + count + '/6</div>' +
        '<div class="trafikljus-label">' + tcLabel + '</div>' +
      '</div>' +
      '<div class="kort">' + provRader + '</div>' +
      '<a class="cta-knapp" href="' + OMPROV_URL + '">Anmäl dig till omprov →</a>' +
      '<div class="info-kort">' +
        '<p class="info-text">För godkänt på kursen krävs 4 av 6 godkända prov. Omprov planeras vid behov — alltid onsdagar kl. 14:45–16:15.</p>' +
      '</div>' +
      '<p class="footer">Historia · Grundskola · Håkan Hildingsson · Vt 2026</p>' +
    '</div>' +
  '</body></html>';
}
