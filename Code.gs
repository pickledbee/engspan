// Backup translator for the English/Español page.
// Uses Google's built-in LanguageApp: free, no API key, nobody signs in.
// Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.

function doGet(e) {
  var q = (e.parameter.q || '').toString();
  var mode = (e.parameter.mode || 'auto').toString();
  var out;
  try {
    if (!q) throw new Error('No text');
    if (mode === 'en') {
      out = { from: 'en', translation: LanguageApp.translate(q, 'en', 'es') };
    } else if (mode === 'es') {
      out = { from: 'es', translation: LanguageApp.translate(q, 'es', 'en') };
    } else {
      // Auto: translate to English. If nothing changed, it was English, so go to Spanish.
      var toEn = LanguageApp.translate(q, '', 'en');
      if (clean(toEn) === clean(q)) {
        out = { from: 'en', translation: LanguageApp.translate(q, 'en', 'es') };
      } else {
        out = { from: 'es', translation: toEn };
      }
    }
  } catch (err) {
    out = { error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function clean(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
}
