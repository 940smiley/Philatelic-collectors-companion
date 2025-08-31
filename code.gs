/***** Philatelic Companion — Google Sheets Setup & Tools (FIXED) *****/
/*
Features:
- One-click setup/repair of sheets, validations, formulas
- Add/Update Gallery images (normalizes Google Drive links to direct-view)
- Reverse image suggestions via Google Cloud Vision (Web Detection)
- Filterable Gallery Views
- Import metadata from Colnect/StampWorld URLs with a Field Mapper sidebar
- Select exactly which fields to write to your log

Notes:
- Set your Google Cloud Vision API Key via: Menu → Configure API Keys (Script Properties preferred).
- A testing fallback is provided via PC_DEFAULTS; remove for production.
*/

const PC = {
  MENU: 'Philatelic Companion',
  SHEETS: {
    LISTS: 'Data – Lists (helper for dropdowns)',
    GALLERY: 'Gallery',
    ASSISTANT: 'Stamp Logging Assistant',
    GALLERY_VIEW: 'Gallery View'
  },
  // Assistant headers (A:V)
  ASSISTANT_HEADERS: [
    'Stamp ID',              // A
    'Country',               // B
    'Year',                  // C
    'Denomination',          // D
    'Condition',             // E
    'Stamp Type',            // F
    'Catalog System',        // G
    'Catalog #',             // H
    'Issue Name',            // I
    'Theme / Topic',         // J
    'For Sale',              // K (checkbox)
    'eBay Listed',           // L (checkbox)
    'Purchase Price',        // M currency
    'Purchase Date',         // N date
    'Source (Dealer)',       // O
    'Market Value',          // P currency
    'Currency',              // Q list
    'Storage Location',      // R list
    'Verified',              // S (checkbox)
    'Notes',                 // T
    'Thumbnail',             // U (formula)
    'Full Image Link'        // V (formula)
  ],
  GALLERY_HEADERS: [
    'Stamp ID',                                   // A
    'Image URL (public or Drive uc?export=view)', // B
    'Thumb URL (optional)',                       // C
    'Caption',                                    // D
    'Notes'                                       // E
  ],
  // Script property keys
  PROP: {
    VISION_API_KEY: 'VISION_API_KEY',
    // Optional (used by Picker sidebar if you choose to wire it)
    PICKER_API_KEY: 'AIzaSyDfbrT955MRTyEh9WRifAXjlBVtKs_f8Ug',
    PICKER_CLIENT_ID: '434272197079-uql8fklmg3ehu1qgdquta5lihlificls.apps.googleusercontent.com'
  },
  // Column indices (1-based) for Assistant
  COL: {
    ID: 1,
    COUNTRY: 2,
    YEAR: 3,
    DENOM: 4,
    CONDITION: 5,
    TYPE: 6,
    CATALOG_SYS: 7,
    CATALOG_NO: 8,
    ISSUE: 9,
    THEME: 10,
    FOR_SALE: 11,
    EBAY: 12,
    BUY_PRICE: 13,
    BUY_DATE: 14,
    SOURCE: 15,
    MARKET_VALUE: 16,
    CURRENCY: 17,
    LOCATION: 18,
    VERIFIED: 19,
    NOTES: 20,
    THUMB: 21,
    LINK: 22
  },
  LAST_ROW: 2000 // prefill formulas down to this row
};

// === Testing defaults (do NOT commit to production) ===
const PC_DEFAULTS = {
  VISION_API_KEY: 'AIzaSyA0EU0myoIDzSl2rg-2D4pmqvpcaUTxtic'
};

/***** Companion extensions & eBay helpers *****/
const PCX = {
  PROP: {
    REQUIRE_IMAGE: 'PC_REQUIRE_IMAGE',
    DEFAULTS: 'PC_DEFAULTS_JSON',           // run presets (country, condition, stampType, etc.)
    IMAGE_SOURCE: 'PC_IMAGE_SOURCE',        // drive|photos|url|none
    PICKER_API_KEY: 'PC_PICKER_API_KEY',
    PICKER_CLIENT_ID: 'PC_PICKER_CLIENT_ID'
  },
  SHEETS: {
    EBAY: 'eBay Listing Template',
    OUTREACH: 'Dealer Outreach Tracker',
    CONSIGN: 'Consignment Form',
    VALUATION: 'Inventory Valuation',
    WISHLIST: 'Wish List',
    INSURANCE: 'Insurance Log',
    ALBUM: 'Album Builder',
    LISTS: 'Data – Lists (helper for dropdowns)',
    DATALISTS: 'Data – Lists (helper for dropdowns)',   // alias used by later helpers
    GALLERY: 'Gallery',
    ASST: 'Stamp Logging Assistant',
    GALLERY_VIEW: 'Gallery View'
  },
  EBAY_HEADER_ROWS_LOCKED: 5
};

/***** Menu *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(PC.MENU)
    .addItem('Setup / Repair Assistant', 'setupStampAssistantWorkbook')
    .addSeparator()
    .addItem('Onboarding Wizard', 'openOnboardingWizard')
    .addItem('Add / Update Gallery Image', 'addOrUpdateGalleryImage')
    .addItem('Add Images via Picker (bulk)', 'openPickerSidebar')
    .addSeparator()
    .addItem('Identify from Image (Vision)', 'identifyFromImage')
    .addItem('Import Metadata from URL', 'importMetadataFromUrl')
    .addSeparator()
    .addItem('Build Gallery View (filters row)', 'buildOrRepairGalleryView')
    .addItem('Rebuild Gallery View (native filters)', 'rebuildGalleryViewWithFilter')
    .addSeparator()
    .addItem('Export selection → eBay (append)', 'exportSelectionToEbay')
    .addSeparator()
    .addItem('Seed Blank Sheets', 'seedAllAuxSheets')
    .addItem('Configure API Keys', 'openConfigSidebar')
    .addSeparator()
    .addItem('Refresh Validations', 'applyAssistantValidations')
    .addItem('Reapply Formulas', 'applyAssistantFormulas')
    .addToUi();

  // Keep DataLists named range current (best-effort; safe no-op if not present)
  try { resyncDataLists(); } catch (_) {}
}

/***** Setup / Repair *****/
function setupStampAssistantWorkbook() {
  const ss = SpreadsheetApp.getActive();
  const dataLists = getOrCreateSheet_(PC.SHEETS.LISTS);
  const gallery = getOrCreateSheet_(PC.SHEETS.GALLERY);
  const assistant = getOrCreateSheet_(PC.SHEETS.ASSISTANT);

  // 1) Seed lists + named ranges
  seedDataLists_(dataLists);
  createNamedRanges_(ss, dataLists);

  // 2) Gallery headers
  if (gallery.getLastRow() < 1) {
    gallery.getRange(1, 1, 1, PC.GALLERY_HEADERS.length).setValues([PC.GALLERY_HEADERS]);
    gallery.setFrozenRows(1);
    gallery.setColumnWidths(1, 5, 200);
  }

  // 3) Assistant headers
  if (assistant.getLastRow() < 1) {
    assistant.getRange(1, 1, 1, PC.ASSISTANT_HEADERS.length).setValues([PC.ASSISTANT_HEADERS]);
    assistant.setFrozenRows(1);
  }

  // 4) Validations, formatting, formulas
  applyAssistantValidations();
  formatAssistant_(assistant);
  applyAssistantFormulas();

  // 5) Build filterable Gallery View
  buildOrRepairGalleryView();

  SpreadsheetApp.getUi().alert(
    'Setup complete.\n\nTips:\n- Add a Gallery row with Stamp ID and image link.\n- Log stamps in the Assistant sheet. Thumbnails & links populate automatically.\n- Use “Identify from Image (Vision)” or “Import Metadata from URL” on a selected Assistant row.'
  );
}

/***** Formulas (Google Sheets–friendly, no fillDown) *****/
function applyAssistantFormulas() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PC.SHEETS.ASSISTANT);
  if (!sh) return;

  const firstRow = 2;
  const lastRow = PC.LAST_ROW;
  const numRows = lastRow - firstRow + 1;

  // A: Stamp ID
  sh.getRange(firstRow, PC.COL.ID, numRows, 1).setFormulaR1C1(
    '=IF(RC2="","", "S"&TEXT(ROW()-1,"00000"))'
  );

  // U: Thumbnail (prefer Gallery thumb else Gallery image)
  sh.getRange(firstRow, PC.COL.THUMB, numRows, 1).setFormulaR1C1(
    '=IFERROR(IMAGE(IF(' +
      'IFNA(VLOOKUP(RC1, Gallery!C[-20]:C[-18], 2, FALSE),"")<>"",' +
      'IFNA(VLOOKUP(RC1, Gallery!C[-20]:C[-18], 2, FALSE),""),' +
      'IFNA(VLOOKUP(RC1, Gallery!C[-20]:C[-19], 2, FALSE),"")' +
    '), 4, 60, 60), "")'
  );

  // V: Full Image Link
  sh.getRange(firstRow, PC.COL.LINK, numRows, 1).setFormulaR1C1(
    '=IFERROR(HYPERLINK(VLOOKUP(RC1, Gallery!C[-21]:C[-20], 2, FALSE), "Open Image"), "")'
  );
}

/***** Validations, formats *****/
function applyAssistantValidations() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PC.SHEETS.ASSISTANT);
  const data = ss.getSheetByName(PC.SHEETS.LISTS);
  if (!sh || !data) return;

  const lastRow = PC.LAST_ROW;
  const named = {
    Countries: ss.getRangeByName('Countries'),
    Denominations: ss.getRangeByName('Denominations'),
    Conditions: ss.getRangeByName('Conditions'),
    StampTypes: ss.getRangeByName('StampTypes'),
    Catalogs: ss.getRangeByName('Catalogs'),
    Themes: ss.getRangeByName('Themes'),
    Currencies: ss.getRangeByName('Currencies'),
    Locations: ss.getRangeByName('Locations')
  };

  const listRule = (rng) => SpreadsheetApp.newDataValidation().requireValueInRange(rng, true).setAllowInvalid(false).build();
  const cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  if (named.Countries)    sh.getRange(2, PC.COL.COUNTRY, lastRow-1).setDataValidation(listRule(named.Countries));
  if (named.Denominations)sh.getRange(2, PC.COL.DENOM, lastRow-1).setDataValidation(listRule(named.Denominations));
  if (named.Conditions)   sh.getRange(2, PC.COL.CONDITION, lastRow-1).setDataValidation(listRule(named.Conditions));
  if (named.StampTypes)   sh.getRange(2, PC.COL.TYPE, lastRow-1).setDataValidation(listRule(named.StampTypes));
  if (named.Catalogs)     sh.getRange(2, PC.COL.CATALOG_SYS, lastRow-1).setDataValidation(listRule(named.Catalogs));
  if (named.Themes)       sh.getRange(2, PC.COL.THEME, lastRow-1).setDataValidation(listRule(named.Themes));
  if (named.Currencies)   sh.getRange(2, PC.COL.CURRENCY, lastRow-1).setDataValidation(listRule(named.Currencies));
  if (named.Locations)    sh.getRange(2, PC.COL.LOCATION, lastRow-1).setDataValidation(listRule(named.Locations));

  sh.getRange(2, PC.COL.FOR_SALE, lastRow-1).setDataValidation(cbRule);
  sh.getRange(2, PC.COL.EBAY, lastRow-1).setDataValidation(cbRule);
  sh.getRange(2, PC.COL.VERIFIED, lastRow-1).setDataValidation(cbRule);

  sh.getRange(2, PC.COL.BUY_DATE, lastRow-1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, PC.COL.BUY_PRICE, lastRow-1).setNumberFormat('"$"#,##0.00');
  sh.getRange(2, PC.COL.MARKET_VALUE, lastRow-1).setNumberFormat('"$"#,##0.00');
}

function formatAssistant_(sh) {
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, PC.ASSISTANT_HEADERS.length);
  sh.getRange(1, 1, 1, PC.ASSISTANT_HEADERS.length).setFontWeight('bold');
  sh.getRange(2, PC.COL.FOR_SALE, PC.LAST_ROW-1).setHorizontalAlignment('center');
  sh.getRange(2, PC.COL.EBAY, PC.LAST_ROW-1).setHorizontalAlignment('center');
  sh.getRange(2, PC.COL.VERIFIED, PC.LAST_ROW-1).setHorizontalAlignment('center');
}

/***** Data lists & named ranges *****/
function seedDataLists_(sh) {
  const lists = {
    Countries: [
      "United States","United Kingdom","Canada","Australia","New Zealand","Germany","France","Italy","Spain","Portugal",
      "Netherlands","Belgium","Switzerland","Austria","Sweden","Norway","Denmark","Finland","Ireland","Poland",
      "Czech Republic","Hungary","Greece","Turkey","Russia","Ukraine","Israel","South Africa","India","China","Japan",
      "South Korea","Hong Kong","Singapore","Taiwan","Thailand","Malaysia","Philippines","Indonesia","Mexico","Brazil",
      "Argentina","Chile","Colombia","Peru","Uruguay","Costa Rica","Monaco","Luxembourg","Iceland"
    ],
    Conditions: ["Mint NH","Mint Hinged","Unused (no gum)","Used","FDC","Cover","Block","Plate","Proof","Specimen"],
    StampTypes: ["Postage","Commemorative","Definitive","Airmail","Postage Due","Revenue","Cinderella","Postal Stationery","Official","Semi-Postal"],
    Catalogs: ["Scott","Michel","Stanley Gibbons","Facit","Yvert & Tellier","Colnect","StampWorld","Unitrade"],
    Themes: ["Animals","Art","Aviation","Birds","Cars","Europa","Flags","Flowers","History","Olympics","Space","Ships","Scouts","Topical"],
    Denominations: ["½","1","2","3","4","5","10","20","25","30","40","50","$0.01","$0.02","$0.03","$0.05","$0.10","$0.20","$0.25","$0.50","$1","$2","$3","$5","$10"],
    Currencies: ["USD","EUR","GBP","JPY","CAD","AUD","CHF","SEK","NOK","DKK","MXN","BRL"],
    Locations: ["Binder A","Binder B","Binder C","Stockbook 1","Stockbook 2","Glassine File 1","Album Page 1","Drawer A","Drawer B","Safe"]
  };

  sh.clear();
  sh.getRange('A1:H1').setValues([['Countries','Conditions','StampTypes','Catalogs','Themes','Denominations','Currencies','Locations']]);
  Object.keys(lists).forEach((key, idx) => {
    const col = idx + 1;
    const values = lists[key].map(v => [v]);
    if (values.length) {
      sh.getRange(2, col, values.length, 1).setValues(values);
      sh.autoResizeColumn(col);
    }
  });
  sh.setFrozenRows(1);
}

function createNamedRanges_(ss, sh) {
  const headers = sh.getRange('A1:H1').getValues()[0];
  headers.forEach((name, i) => {
    const col = i + 1;
    const last = sh.getRange(sh.getMaxRows(), col).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    const rng = sh.getRange(2, col, Math.max(0, last - 1), 1);
    if (rng.getNumRows() > 0) ss.setNamedRange(name, rng);
  });
}

/***** Helpers *****/
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function normalizeDriveUrl_(url) {
  try {
    const fileIdMatch = url.match(/[-\w]{25,}/);
    if (fileIdMatch) {
      return `https://drive.google.com/uc?export=view&id=${fileIdMatch[0]}`;
    }
    return url;
  } catch (e) {
    return url;
  }
}

/***** Add / Update Gallery Image *****/
function addOrUpdateGalleryImage() {
  const ui = SpreadsheetApp.getUi();
  const idResp = ui.prompt('Add / Update Gallery', 'Enter Stamp ID (e.g., S00001):', ui.ButtonSet.OK_CANCEL);
  if (idResp.getSelectedButton() !== ui.Button.OK) return;
  const stampId = idResp.getResponseText().trim();
  if (!stampId) return;

  const imgResp = ui.prompt('Image URL or Google Drive link', 'Paste a public image URL or a Drive link (we will normalize it):', ui.ButtonSet.OK_CANCEL);
  if (imgResp.getSelectedButton() !== ui.Button.OK) return;
  const imgUrl = normalizeDriveUrl_(imgResp.getResponseText().trim());
  if (!imgUrl) return;

  const thumbResp = ui.prompt('Optional thumbnail URL', 'Leave blank to auto-use full image as thumbnail.', ui.ButtonSet.OK_CANCEL);
  if (thumbResp.getSelectedButton() !== ui.Button.OK) return;
  const thumbUrl = thumbResp.getResponseText().trim() ? normalizeDriveUrl_(thumbResp.getResponseText().trim()) : '';

  const capResp = ui.prompt('Caption (optional)', 'Enter an optional caption:', ui.ButtonSet.OK_CANCEL);
  if (capResp.getSelectedButton() !== ui.Button.OK) return;
  const caption = capResp.getResponseText().trim();

  const sh = SpreadsheetApp.getActive().getSheetByName(PC.SHEETS.GALLERY);
  const ids = sh.getRange(2, 1, Math.max(0, sh.getLastRow()-1), 1).getDisplayValues().map(r => r[0]);
  let row = ids.indexOf(stampId);
  if (row === -1) {
    sh.appendRow([stampId, imgUrl, thumbUrl, caption, '']);
  } else {
    const r = row + 2;
    sh.getRange(r, 1, 1, 5).setValues([[stampId, imgUrl, thumbUrl, caption, '']]);
  }
  SpreadsheetApp.getUi().alert('Gallery updated.');
}

/***** Identify from Image (Google Cloud Vision – Web Detection) *****/
function identifyFromImage() {
  const ss = SpreadsheetApp.getActive();
  const asst = ss.getSheetByName(PC.SHEETS.ASSISTANT);
  if (!asst) return;

  const row = asst.getActiveCell() ? asst.getActiveCell().getRow() : 0;
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Select a data row in "Stamp Logging Assistant" first.');
    return;
  }

  const stampId = asst.getRange(row, PC.COL.ID).getDisplayValue();
  if (!stampId) {
    SpreadsheetApp.getUi().alert('This row has no Stamp ID yet. Fill Country first to generate an ID.');
    return;
  }

  const imgUrl = findImageUrlById_(stampId);
  if (!imgUrl) {
    SpreadsheetApp.getUi().alert('No image found in Gallery for this Stamp ID. Add one first.');
    return;
  }

  const suggestions = getVisionWebSuggestions_(imgUrl) || {};
  const candidateUrls = (suggestions.pages || []).slice(0, 8);

  const prefill = {
    row: row,
    stampId: stampId,
    country: suggestions.bestGuessCountry || '',
    year: suggestions.bestGuessYear || '',
    theme: suggestions.bestGuessTheme || '',
    issue: suggestions.bestGuessIssue || '',
    catalogSystem: '',
    catalogNumber: '',
    denomination: ''
  };

  openFieldMapperSidebar_({
    title: 'Reverse Image Suggestions',
    subtitle: 'Select which suggestions to apply to the current row.',
    fields: prefill,
    candidates: candidateUrls
  });
}

function findImageUrlById_(stampId) {
  const sh = SpreadsheetApp.getActive().getSheetByName(PC.SHEETS.GALLERY);
  if (!sh) return '';
  const last = sh.getLastRow();
  if (last < 2) return '';
  const values = sh.getRange(2, 1, last - 1, 3).getValues(); // A:ID, B:Image, C:Thumb
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === stampId) {
      return values[i][1] || values[i][2] || '';
    }
  }
  return '';
}

function getVisionWebSuggestions_(imageUrl) {
  const key = PropertiesService.getScriptProperties().getProperty(PC.PROP.VISION_API_KEY) || PC_DEFAULTS.VISION_API_KEY;
  if (!key) {
    SpreadsheetApp.getUi().alert('Set your Google Cloud Vision API key in “Configure API Keys”.');
    return {};
  }
  const endpoint = 'https://vision.googleapis.com/v1/images:annotate?key=' + encodeURIComponent(key);
  const payload = {
    requests: [{
      image: { source: { imageUri: imageUrl } },
      features: [{ type: 'WEB_DETECTION' }]
    }]
  };
  const res = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const data = JSON.parse(res.getContentText());
  const web = (((data || {}).responses || [])[0] || {}).webDetection || {};

  // Heuristics to guess country/year/theme from entities or bestGuessLabels
  const bestGuess = (web.bestGuessLabels && web.bestGuessLabels[0] && web.bestGuessLabels[0].label) || '';
  const entities = (web.webEntities || []).map(e => e.description || '');
  const allText = (bestGuess + ' ' + entities.join(' ')).toLowerCase();
  const yearMatch = allText.match(/\b(18|19|20)\d{2}\b/);
  const country = guessCountryFromText_(allText);
  const theme = guessThemeFromText_(allText);

  return {
    bestGuess: bestGuess,
    entities: entities,
    pages: (web.pagesWithMatchingImages || []).map(p => p.url).filter(Boolean),
    bestGuessYear: yearMatch ? yearMatch[0] : '',
    bestGuessCountry: country || '',
    bestGuessTheme: theme || '',
    bestGuessIssue: ''
  };
}

function guessCountryFromText_(txt) {
  const countries = SpreadsheetApp.getActive().getRangeByName('Countries');
  if (!countries) return '';
  const vals = countries.getDisplayValues().flat().filter(Boolean);
  for (const c of vals) {
    if (txt.includes(String(c).toLowerCase())) return c;
  }
  return '';
}

function guessThemeFromText_(txt) {
  const themes = SpreadsheetApp.getActive().getRangeByName('Themes');
  if (!themes) return '';
  const vals = themes.getDisplayValues().flat().filter(Boolean);
  for (const t of vals) {
    if (txt.includes(String(t).toLowerCase())) return t;
  }
  return '';
}

/***** Import Metadata from URL (Colnect, StampWorld) with selective mapping *****/
function importMetadataFromUrl() {
  const ui = SpreadsheetApp.getUi();
  const urlResp = ui.prompt('Import Metadata', 'Paste a Colnect or StampWorld item URL:', ui.ButtonSet.OK_CANCEL);
  if (urlResp.getSelectedButton() !== ui.Button.OK) return;
  const targetUrl = urlResp.getResponseText().trim();
  if (!targetUrl) return;

  const asst = SpreadsheetApp.getActive().getSheetByName(PC.SHEETS.ASSISTANT);
  const row = asst.getActiveCell() ? asst.getActiveCell().getRow() : 0;
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Select a data row in "Stamp Logging Assistant" first.');
    return;
  }

  const parsed = parseStampPage_(targetUrl);
  if (!parsed || Object.keys(parsed).length === 0) {
    SpreadsheetApp.getUi().alert('Could not extract fields from that URL. You can still use the Field Mapper to manually enter values.');
  }

  openFieldMapperSidebar_({
    title: 'Import Metadata',
    subtitle: 'Choose which fields to import into the selected row.',
    fields: Object.assign({
      row: row,
      stampId: asst.getRange(row, PC.COL.ID).getDisplayValue() || ''
    }, parsed),
    candidates: []
  });
}

function parseStampPage_(url) {
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() >= 400) return {};
    const html = resp.getContentText();

    if (/colnect\.com/i.test(url)) {
      return parseColnect_(html);
    } else if (/stampworld\./i.test(url)) {
      return parseStampWorld_(html);
    } else {
      return parseGenericStampPage_(html);
    }
  } catch (e) {
    return {};
  }
}

function parseColnect_(html) {
  const get = (label) => {
    const re = new RegExp(label + '\\s*</[^>]+>\\s*<[^>]+>(.*?)<', 'i');
    const m = html.match(re);
    return m ? sanitizeHtmlText_(m[1]) : '';
  };
  const title = metaContent_(html, 'og:title') || '';
  const desc = metaContent_(html, 'og:description') || '';

  const country = get('Country') || firstBeforeDash_(title);
  const year = firstYear_(html) || firstYear_(title) || firstYear_(desc);
  const denom = get('Face value') || get('Denomination');
  const issue = get('Series') || '';

  const catLine = (html.match(/Catalog(?:\s+codes|\s*#?).*?<\/[^>]+>\s*<[^>]+>(.*?)</i) || [])[1] || '';
  let catalogSystem = '', catalogNumber = '';
  if (catLine) {
    const sc = catLine.match(/Scott:\s*([^;]+)/i);
    const mi = catLine.match(/Mi:\s*([^;]+)/i);
    if (sc) { catalogSystem = 'Scott'; catalogNumber = sanitizeHtmlText_(sc[1]); }
    else if (mi) { catalogSystem = 'Michel'; catalogNumber = sanitizeHtmlText_(mi[1]); }
  }
  const theme = get('Topic') || get('Theme') || metaContent_(html, 'keywords');

  return {
    urlSource: 'Colnect',
    sourceUrl: '',
    country: country,
    year: year,
    denomination: denom,
    issue: issue,
    catalogSystem: catalogSystem,
    catalogNumber: catalogNumber,
    theme: theme
  };
}

function parseStampWorld_(html) {
  const title = metaContent_(html, 'og:title') || '';
  const desc = metaContent_(html, 'og:description') || '';

  const country = firstBeforeDash_(title);
  const year = firstYear_(html) || firstYear_(title) || firstYear_(desc);
  const denom = findAfter_(html, /(?:Value|Denomination):\s*<\/td>\s*<td[^>]*>(.*?)<\/td>/i);
  const issue = findAfter_(html, /(?:Issue|Series):\s*<\/td>\s*<td[^>]*>(.*?)<\/td>/i);

  let catalogSystem = '';
  let catalogNumber = '';
  const catLine = findAfter_(html, /Catalog\s*number.*?<\/td>\s*<td[^>]*>(.*?)<\/td>/i);
  if (catLine) {
    const sc = catLine.match(/Scott:\s*([^;]+)/i);
    const mi = catLine.match(/Mi:\s*([^;]+)/i);
    const sg = catLine.match(/SG:\s*([^;]+)/i);
    if (sc) { catalogSystem = 'Scott'; catalogNumber = sanitizeHtmlText_(sc[1]); }
    else if (mi) { catalogSystem = 'Michel'; catalogNumber = sanitizeHtmlText_(mi[1]); }
    else if (sg) { catalogSystem = 'Stanley Gibbons'; catalogNumber = sanitizeHtmlText_(sg[1]); }
  }

  const theme = findAfter_(html, /(?:Theme|Topic):\s*<\/td>\s*<td[^>]*>(.*?)<\/td>/i)
             || metaContent_(html, 'keywords');

  return {
    urlSource: 'StampWorld',
    sourceUrl: '',
    country: country,
    year: year,
    denomination: sanitizeHtmlText_(denom),
    issue: sanitizeHtmlText_(issue),
    catalogSystem: catalogSystem,
    catalogNumber: catalogNumber,
    theme: sanitizeHtmlText_(theme)
  };
}

// NEW: generic fallback for non-Colnect/StampWorld pages
function parseGenericStampPage_(html) {
  const title = metaContent_(html, 'og:title') || metaContent_(html, 'twitter:title') || '';
  const desc = metaContent_(html, 'og:description') || metaContent_(html, 'description') || '';
  const keywords = metaContent_(html, 'keywords') || '';
  const blob = [title, desc, keywords].join(' ');
  const year = firstYear_((blob));
  const country = firstBeforeDash_(title);
  return {
    urlSource: 'Generic',
    sourceUrl: '',
    country: country || '',
    year: year || '',
    denomination: '',
    issue: '',
    catalogSystem: '',
    catalogNumber: '',
    theme: sanitizeHtmlText_(keywords || '')
  };
}

/***** Field Mapper UI *****/
function openFieldMapperSidebar_(payload) {
  const safe = JSON.stringify(payload || {});
  const html = HtmlService.createHtmlOutput(buildFieldMapperHtml_(safe)).setTitle('Field Mapper');
  SpreadsheetApp.getUi().showSidebar(html);
}

function buildFieldMapperHtml_(payloadJson) {
  return `
  <div style="font: 13px/1.4 Roboto,Arial; padding:12px 14px;">
    <h3 style="margin:0 0 4px;">Field Mapper</h3>
    <div id="sub" style="color:#666; margin-bottom:8px;"></div>
    <form id="f" onsubmit="return false;"></form>
    <div id="candidates" style="margin-top:10px;"></div>
    <div style="margin-top:10px; display:flex; gap:8px;">
      <button onclick="apply()" style="padding:6px 10px;">Apply Selected</button>
      <button onclick="google.script.host.close()" style="padding:6px 10px;">Close</button>
    </div>
    <script>
      const payload = ${payloadJson};
      document.getElementById('sub').textContent = payload.subtitle || '';
      const f = document.getElementById('f');
      const fields = payload.fields || {};
      const order = ['country','year','denomination','issue','catalogSystem','catalogNumber','theme'];
      order.forEach(k=>{
        const val = fields[k] || '';
        if (val === '') return;
        const id = 'fld_'+k;
        const row = document.createElement('div');
        row.style.margin = '6px 0';
        row.innerHTML = '<label><input type="checkbox" id="'+id+'" checked> <b>'+k+':</b> '+String(val)+'</label>';
        f.appendChild(row);
      });
      if ((payload.candidates||[]).length){
        const box = document.getElementById('candidates');
        box.innerHTML = '<div style="margin:6px 0 4px; font-weight:bold;">Candidate pages</div>' +
          payload.candidates.map(u => '<div style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;"><a target="_blank" href="'+u+'">'+u+'</a></div>').join('');
      }
      function apply(){
        const out = {};
        (order).forEach(k=>{
          const val = fields[k] || '';
          const el = document.getElementById('fld_'+k);
          if (val && el && el.checked) out[k] = val;
        });
        google.script.run.withSuccessHandler(()=>google.script.host.close())
          .applyFieldMapping_(fields.row || 0, out);
      }
      window.apply = apply;
    </script>
  </div>`;
}

function applyFieldMapping_(row, selected) {
  if (!row || !selected) return;
  const sh = SpreadsheetApp.getActive().getSheetByName(PC.SHEETS.ASSISTANT);
  if (!sh) return;

  const map = {
    country: PC.COL.COUNTRY,
    year: PC.COL.YEAR,
    denomination: PC.COL.DENOM,
    issue: PC.COL.ISSUE,
    catalogSystem: PC.COL.CATALOG_SYS,
    catalogNumber: PC.COL.CATALOG_NO,
    theme: PC.COL.THEME
  };
  Object.keys(selected).forEach(key => {
    const col = map[key];
    if (col) sh.getRange(row, col).setValue(String(selected[key]));
  });
  sh.getRange(row, PC.COL.VERIFIED).setValue(true);
}

/***** HTML parsing helpers *****/
function metaContent_(html, name) {
  if (!html || !name) return '';
  const n = name.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
  const re1 = new RegExp('<meta[^>]+(?:property|name)=[\\"\\\']' + n + '[\\"\\\'][^>]+content=[\\"\\\']([^"\\\']+)[\\"\\\']', 'i');
  const re2 = new RegExp('<meta[^>]+content=[\\"\\\']([^"\\\']+)[\\"\\\'][^>]+(?:property|name)=[\\"\\\']' + n + '[\\"\\\']', 'i');
  const m = html.match(re1) || html.match(re2);
  return m ? sanitizeHtmlText_(m[1]) : '';
}

function sanitizeHtmlText_(s) {
  if (!s) return '';
  let t = String(s).replace(/<[^>]*>/g, ' ');
  t = t.replace(/&nbsp;/g, ' ')
       .replace(/&amp;/g, '&')
       .replace(/&lt;/g, '<')
       .replace(/&gt;/g, '>')
       .replace(/&quot;/g, '"')
       .replace(/&#39;/g, "'");
  return t.replace(/\s+/g, ' ').trim();
}

function firstYear_(s) {
  if (!s) return '';
  const m = String(s).match(/\b(18|19|20)\d{2}\b/);
  return m ? m[0] : '';
}

function firstBeforeDash_(s) {
  if (!s) return '';
  const parts = String(s).split(' - ');
  return parts.length ? parts[0].trim() : '';
}

function findAfter_(html, re) {
  if (!html || !re) return '';
  const m = html.match(re);
  return m ? sanitizeHtmlText_(m[1]) : '';
}

/***** Seeders & DataLists (aux sheets) *****/
function seedAllAuxSheets() {
  const ss = SpreadsheetApp.getActive();

  const toSeed = [
    { name: PCX.SHEETS.OUTREACH, headers: ['Dealer', 'Contact Info', 'Notes', 'Status', 'Follow Up Date'] },
    { name: PCX.SHEETS.CONSIGN, headers: ['Consignment ID', 'Item', 'Value', 'Split %', 'Consignor', 'Notes'] },
    { name: PCX.SHEETS.VALUATION, headers: ['Date', 'Total Value', 'Method', 'Notes'] },
    { name: PCX.SHEETS.WISHLIST, headers: ['Country', 'Year', 'Denomination', 'Condition', 'Catalog #', 'Notes'] },
    { name: PCX.SHEETS.INSURANCE, headers: ['Policy #', 'Provider', 'Coverage Amount', 'Premium', 'Expiry Date', 'Notes'] },
    { name: PCX.SHEETS.ALBUM, headers: ['Album Name', 'Page #', 'Stamp ID', 'Notes'] },
    { name: PCX.SHEETS.DATALISTS, headers: ['List Name', 'Value'] }
  ];

  toSeed.forEach(cfg => {
    let sh = ss.getSheetByName(cfg.name);
    if (!sh) sh = ss.insertSheet(cfg.name);
    else sh.clear();
    if (cfg.headers?.length) {
      sh.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]);
      sh.setFrozenRows(1);
    }
  });

  const dataListSheet = ss.getSheetByName(PCX.SHEETS.DATALISTS);
  const dataLists = {
    'Denominations': [
      '1¢', '2¢', '3¢', '4¢', '5¢', '6¢', '7¢', '8¢', '9¢', '10¢', '12¢', '15¢',
      '20¢', '25¢', '30¢', '50¢', '$1', '$2', '$3', '$4', '$5', '$10', '$20', '$50', '$100'
    ],
    'Themes': [
      'Animals', 'Birds', 'Insects', 'Marine Life', 'Plants', 'Sports', 'Historical Events',
      'Space Exploration', 'Art & Culture', 'Architecture', 'Transportation', 'Maps',
      'Flags', 'Holidays', 'Scouts', 'Olympics'
    ],
    'Conditions': [
      'Mint', 'Mint Hinged', 'Mint No Gum', 'Used', 'CTO', 'Damaged'
    ]
  };

  let row = 2;
  for (let [listName, values] of Object.entries(dataLists)) {
    values.forEach(val => {
      dataListSheet.getRange(row, 1).setValue(listName);
      dataListSheet.getRange(row, 2).setValue(val);
      row++;
    });
  }

  SpreadsheetApp.getUi().alert('Auxiliary sheets and DataLists have been created/seeded.');
}

function resyncDataLists() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(PCX.SHEETS.DATALISTS);
  if (!sheet) return;

  const existing = sheet.getRange(2, 1, Math.max(0, sheet.getLastRow() - 1), 2)
    .getValues()
    .map(r => ({ list: r[0], value: r[1] }));

  const dataLists = {
    'Denominations': [
      '1¢', '2¢', '3¢', '4¢', '5¢', '6¢', '7¢', '8¢', '9¢', '10¢', '12¢', '15¢',
      '20¢', '25¢', '30¢', '50¢', '$1', '$2', '$3', '$4', '$5', '$10', '$20', '$50', '$100'
    ],
    'Themes': [
      'Animals', 'Birds', 'Insects', 'Marine Life', 'Plants', 'Sports', 'Historical Events',
      'Space Exploration', 'Art & Culture', 'Architecture', 'Transportation', 'Maps',
      'Flags', 'Holidays', 'Scouts', 'Olympics'
    ],
    'Conditions': [
      'Mint', 'Mint Hinged', 'Mint No Gum', 'Used', 'CTO', 'Damaged'
    ]
  };

  let row = sheet.getLastRow() + 1;
  for (let [listName, values] of Object.entries(dataLists)) {
    values.forEach(val => {
      if (!existing.find(e => e.list === listName && e.value === val)) {
        sheet.getRange(row, 1).setValue(listName);
        sheet.getRange(row, 2).setValue(val);
        row++;
      }
    });
  }
}

/***** Onboarding wizard *****/
function openOnboardingWizard() {
  const props = PropertiesService.getScriptProperties();
  const prior = {
    requireImage: props.getProperty(PCX.PROP.REQUIRE_IMAGE) === 'true',
    imageSource: props.getProperty(PCX.PROP.IMAGE_SOURCE) || 'drive',
    defaults: JSON.parse(props.getProperty(PCX.PROP.DEFAULTS) || '{}')
  };
  const html = HtmlService.createHtmlOutput(`
    <div style="font:13px/1.5 Roboto,Arial; padding:12px; width:360px">
      <h3 style="margin:0 0 8px;">Onboarding</h3>
      <p>Answer a few questions to streamline this run. You can change any setting later.</p>

      <label style="display:block; margin-top:8px;">How do you store your stamps?</label>
      <select id="locations" style="width:100%">
        <option>Binders & Album Pages</option>
        <option>Stockbooks</option>
        <option>Glassines / Envelopes</option>
        <option>Drawers / Safe</option>
        <option>Custom...</option>
      </select>
      <input id="locationsCustom" placeholder="Custom locations (comma-separated)" style="width:100%; margin-top:6px; display:none"/>

      <label style="display:block; margin-top:12px;">Will you be using photos?</label>
      <select id="usePhotos" style="width:100%">
        <option value="yes">Yes</option>
        <option value="no">No</option>
      </select>

      <div id="photoBlock" style="margin-top:8px;">
        <label>Where are the photos?</label>
        <select id="imageSource" style="width:100%">
          <option value="drive" ${prior.imageSource==='drive'?'selected':''}>Google Drive (recommended)</option>
          <option value="photos" ${prior.imageSource==='photos'?'selected':''}>Google Photos (Picker)</option>
          <option value="url" ${prior.imageSource==='url'?'selected':''}>Paste image URLs</option>
          <option value="none" ${prior.imageSource==='none'?'selected':''}>Not using images</option>
        </select>
        <label style="display:block; margin-top:8px;">
          <input type="checkbox" id="requireImage" ${prior.requireImage?'checked':''}/> Require an image for each stamp
        </label>
      </div>

      <div style="margin-top:12px;">
        <label>Will you use image search to identify stamps?</label>
        <select id="useVision" style="width:100%">
          <option value="yes">Yes</option>
          <option value="no">No</option>
        </select>
      </div>

      <div style="margin-top:12px;">
        <label>Any file imports to begin?</label>
        <textarea id="imports" placeholder="Paste CSV/Sheet URLs (one per line)" style="width:100%; min-height:64px;"></textarea>
      </div>

      <div style="margin-top:12px;">
        <label>Preset attributes for this run (optional):</label>
        <input id="defCountry" placeholder="Country" style="width:100%; margin-top:6px"/>
        <input id="defCondition" placeholder="Condition (e.g., Mint NH)" style="width:100%; margin-top:6px"/>
        <input id="defType" placeholder="Stamp Type (e.g., Commemorative)" style="width:100%; margin-top:6px"/>
        <input id="defTheme" placeholder="Theme/Topic" style="width:100%; margin-top:6px"/>
      </div>

      <div style="margin-top:14px; display:flex; gap:8px;">
        <button onclick="save()" style="padding:6px 10px;">Save</button>
        <button onclick="bulk()" style="padding:6px 10px;">Bulk add images now</button>
        <button onclick="google.script.host.close()" style="padding:6px 10px;">Close</button>
      </div>

      <script>
        const locSel = document.getElementById('locations');
        const locCustom = document.getElementById('locationsCustom');
        locSel.addEventListener('change', () => {
          locCustom.style.display = (locSel.value === 'Custom...') ? 'block' : 'none';
        });
        function gather(){
          return {
            locationsChoice: locSel.value,
            locationsCustom: locCustom.value,
            usePhotos: document.getElementById('usePhotos').value,
            imageSource: document.getElementById('imageSource').value,
            requireImage: document.getElementById('requireImage').checked,
            useVision: document.getElementById('useVision').value,
            imports: document.getElementById('imports').value,
            defaults: {
              country: document.getElementById('defCountry').value,
              condition: document.getElementById('defCondition').value,
              stampType: document.getElementById('defType').value,
              theme: document.getElementById('defTheme').value
            }
          };
        }
        function save(){
          google.script.run.withSuccessHandler(()=>google.script.host.close())
            .applyOnboardingSettings_(gather());
        }
        function bulk(){
          google.script.run.withSuccessHandler(()=>google.script.host.close())
            .applyOnboardingSettings_(gather());
          google.script.run.openPickerSidebar();
        }
        window.save = save; window.bulk = bulk;
      </script>
    </div>
  `).setTitle('Onboarding');
  SpreadsheetApp.getUi().showSidebar(html);
}

function applyOnboardingSettings_(form) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PCX.PROP.REQUIRE_IMAGE, String(!!form.requireImage));
  props.setProperty(PCX.PROP.IMAGE_SOURCE, form.imageSource || 'drive');
  props.setProperty(PCX.PROP.DEFAULTS, JSON.stringify(form.defaults || {}));

  // Expand Data Lists with custom locations
  const lists = SpreadsheetApp.getActive().getSheetByName(PC.SHEETS.LISTS || PCX.SHEETS.LISTS);
  if (lists && form.locationsChoice) {
    const colH = 8; // H = Locations
    let values = [];
    if (form.locationsChoice === 'Custom...' && form.locationsCustom) {
      values = String(form.locationsCustom).split(',').map(s => [s.trim()]).filter(r => r[0]);
    } else {
      const preset = {
        'Binders & Album Pages': ['Binder A','Binder B','Album Page 1','Album Page 2','Slipcase A','Slipcase B'],
        'Stockbooks': ['Stockbook 1','Stockbook 2','Stockbook 3','Stockbook 4'],
        'Glassines / Envelopes': ['Glassine File 1','Glassine File 2','Envelope Bin A','Envelope Bin B'],
        'Drawers / Safe': ['Drawer A','Drawer B','Safe']
      }[form.locationsChoice] || [];
      values = preset.map(s => [s]);
    }
    if (values.length) {
      lists.getRange(2, colH, lists.getMaxRows()-1, 1).clearContent();
      lists.getRange(2, colH, values.length, 1).setValues(values);
      SpreadsheetApp.getActive().setNamedRange('Locations', lists.getRange(2, colH, values.length, 1));
    }
  }
}

/***** Picker sidebar (Drive + Photos) for bulk images *****/
function openPickerSidebar() {
  const props = PropertiesService.getScriptProperties();
  // ✅ Fallback to constants if properties are blank
  const KEY = props.getProperty(PCX.PROP.PICKER_API_KEY) || PC.PROP.PICKER_API_KEY || '';
  const CID = props.getProperty(PCX.PROP.PICKER_CLIENT_ID) || PC.PROP.PICKER_CLIENT_ID || '';

  const html = HtmlService.createHtmlOutput(`
    <div style="font:13px/1.5 Roboto,Arial; padding:12px; width:360px">
      <h3 style="margin:0 0 8px;">Add images from Drive / Photos</h3>
      <p>Select multiple images. We’ll create Gallery rows and optional Assistant rows with auto-generated Stamp IDs.</p>

      <div style="margin-top:8px">
        <button id="pickDrive" style="padding:6px 10px" disabled>Pick from Drive</button>
        <button id="pickPhotos" style="padding:6px 10px" disabled>Pick from Google Photos</button>
      </div>

      <div id="log" style="margin-top:10px; color:#444; white-space:pre-wrap;"></div>

      <script src="https://accounts.google.com/gsi/client" async defer></script>
      <script src="https://apis.google.com/js/api.js"></script>
      <script>
        const API_KEY = ${JSON.stringify(KEY)};
        const CLIENT_ID = ${JSON.stringify(CID)};
        const SCOPES = 'https://www.googleapis.com/auth/drive.readonly https://www.googleapis.com/auth/photoslibrary.readonly';
        const logEl = document.getElementById('log');
        const btnDrive = document.getElementById('pickDrive');
        const btnPhotos = document.getElementById('pickPhotos');
        let token = null, pickerReady = false;

        function log(msg){ logEl.textContent += (msg + '\\n'); }

        // Basic preflight
        if (!API_KEY || !CLIENT_ID){
          log('❌ Missing Picker credentials. Set them via “Configure API Keys” or call setPickerDefaults_().');
        }

        // Load Picker library
        function initPicker(){
          try {
            gapi.load('picker', () => {
              pickerReady = true;
              if (API_KEY && CLIENT_ID) enableButtons();
              log('✅ Picker library loaded.');
            });
          } catch (e) {
            log('❌ Failed to load Picker: ' + e);
          }
        }

        function enableButtons(){
          btnDrive.disabled = false;
          btnPhotos.disabled = false;
        }

        // OAuth token via GIS
        function ensureToken(cb){
          if (!API_KEY || !CLIENT_ID){
            log('❌ API key / Client ID not configured.');
            return;
          }
          if (token) return cb(token);
          try {
            const tc = google.accounts.oauth2.initTokenClient({
              client_id: CLIENT_ID,
              scope: SCOPES,
              callback: (t) => { token = t.access_token; cb(token); }
            });
            tc.requestAccessToken({ prompt: '' });
          } catch(e){
            log('❌ OAuth init failed: ' + (e && e.message ? e.message : e));
          }
        }

        function pickDrive(){
          if (!pickerReady) return log('⏳ Picker not ready.');
          ensureToken(function(tok){
            try {
              const picker = new google.picker.PickerBuilder()
                .addView(google.picker.ViewId.DOCS_IMAGES)
                .enableFeature(google.picker.Feature.MULTISELECT_ENABLED)
                .setOAuthToken(tok)
                .setDeveloperKey(API_KEY)
                .setCallback(function(data){
                  if (data.action === google.picker.Action.PICKED){
                    const items = data.docs.map(d => ({name:d.name, url:d.url, id:d.id, mime:d.mimeType}));
                    google.script.run.withSuccessHandler(msg=>log('✅ ' + msg))
                      .withFailureHandler(err=>log('❌ Server error: ' + err))
                      .bulkIngestPickedImages_(items, 'drive');
                  } else if (data.action === google.picker.Action.CANCEL){
                    log('ℹ️ Drive picker canceled.');
                  }
                })
                .build();
              picker.setVisible(true);
            } catch(e){
              log('❌ Drive picker failed: ' + e);
            }
          });
        }

        function pickPhotos(){
          if (!pickerReady) return log('⏳ Picker not ready.');
          ensureToken(function(tok){
            try {
              const view = new google.picker.View(google.picker.ViewId.PHOTOS);
              const picker = new google.picker.PickerBuilder()
                .addView(view)
                .enableFeature(google.picker.Feature.MULTISELECT_ENABLED)
                .setOAuthToken(tok)
                .setDeveloperKey(API_KEY)
                .setCallback(function(data){
                  if (data.action === google.picker.Action.PICKED){
                    const items = data.docs.map(d => ({name:d.name, url:d.url || (d.thumbnails&&d.thumbnails[0]&&d.thumbnails[0].url) || '', id:d.id, mime:d.mimeType}));
                    google.script.run.withSuccessHandler(msg=>log('✅ ' + msg))
                      .withFailureHandler(err=>log('❌ Server error: ' + err))
                      .bulkIngestPickedImages_(items, 'photos');
                  } else if (data.action === google.picker.Action.CANCEL){
                    log('ℹ️ Photos picker canceled.');
                  }
                })
                .build();
              picker.setVisible(true);
            } catch(e){
              log('❌ Photos picker failed: ' + e);
            }
          });
        }

        // Wire buttons after libraries are ready
        window.addEventListener('load', () => {
          try {
            initPicker();
            btnDrive.onclick = pickDrive;
            btnPhotos.onclick = pickPhotos;
          } catch(e){
            log('❌ Init error: ' + e);
          }
        });
      </script>
    </div>
  `).setTitle('Image Picker');
  SpreadsheetApp.getUi().showSidebar(html);
}


function bulkIngestPickedImages_(items, source) {
  const gal = SpreadsheetApp.getActive().getSheetByName(PCX.SHEETS.GALLERY);
  const asst = SpreadsheetApp.getActive().getSheetByName(PCX.SHEETS.ASST);
  if (!gal || !asst || !items || !items.length) return 'No items selected.';

  let created = 0;

  items.forEach((it) => {
    let url = it.url || '';
    if (source === 'drive' && it.id) {
      url = `https://drive.google.com/uc?export=view&id=${it.id}`;
    }
    if (!url) return;

    const row = asst.getLastRow() < 2 ? 2 : asst.getLastRow() + 1;
    const stampId = 'S' + Utilities.formatString('%05d', row - 1);

    gal.appendRow([stampId, url, '', it.name || '', '']);

    const props = PropertiesService.getScriptProperties();
    const def = JSON.parse(props.getProperty(PCX.PROP.DEFAULTS) || '{}');
    const values = Array(22).fill('');
    values[0] = stampId;                 // A Stamp ID
    values[1] = def.country || '';       // B Country
    values[4] = def.condition || '';     // E Condition
    values[5] = def.stampType || '';     // F Stamp Type
    values[9] = def.theme || '';         // J Theme
    asst.appendRow(values);
    created++;
  });

  return `Added ${created} image(s). Assistant rows and Gallery updated.`;
}

/***** Require-image enforcement and auto-ID assist *****/
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== PC.SHEETS.ASSISTANT && sh.getName() !== PCX.SHEETS.ASST) return;

    const props = PropertiesService.getScriptProperties();
    const requireImage = props.getProperty(PCX.PROP.REQUIRE_IMAGE) === 'true';
    const r = e.range.getRow();
    if (r < 2) return;

    const id = sh.getRange(r, 1).getDisplayValue(); // Stamp ID
    if (requireImage) {
      const hasImage = hasGalleryImage_(id);
      sh.getRange(r, 1, 1, 22).setBackground(hasImage ? null : '#fff2f2');
      if (!hasImage) {
        SpreadsheetApp.getActive().toast('Image required — add via Gallery or Picker, then continue.', 'Missing image', 4);
      }
    }
  } catch (_) {}
}
function hasGalleryImage_(stampId) {
  if (!stampId) return false;
  const gal = SpreadsheetApp.getActive().getSheetByName(PCX.SHEETS.GALLERY);
  if (!gal) return false;
  const last = gal.getLastRow();
  if (last < 2) return false;
  const vals = gal.getRange(2, 1, last-1, 2).getValues(); // A:ID, B:URL
  return vals.some(r => r[0] === stampId && r[1]);
}

/***** Gallery Views *****/
function buildOrRepairGalleryView() {
  const ss = SpreadsheetApp.getActive();
  const galName = PC.SHEETS.GALLERY;
  let sh = ss.getSheetByName(PC.SHEETS.GALLERY_VIEW);
  if (!sh) sh = ss.insertSheet(PC.SHEETS.GALLERY_VIEW);
  sh.clear();

  // Filter controls row
  sh.getRange('A1:J1').setValues([[
    'Filters', 'Country', '', 'Year From', 'Year To', '', 'Denomination', '', 'Theme', ''
  ]]);
  sh.getRange('A1:J1').setFontWeight('bold');
  sh.getRange('D2:E2').setNumberFormat('0');
  sh.setFrozenRows(3);

  // Headers
  sh.getRange('A4:J4').setValues([[
    'Stamp ID','Country','Year','Denomination','Theme','Issue','Condition','Thumbnail','Open Image','Caption'
  ]]).setFontWeight('bold');

  const formula = `=ARRAYFORMULA(
    QUERY(
      {
        'Stamp Logging Assistant'!A:V,
        IFERROR(VLOOKUP('Stamp Logging Assistant'!A:A, ${galName}!A:E, 4, FALSE), "")
      },
      "select Col1,Col2,Col3,Col4,Col10,Col9,Col5,Col21,Col22,Col26 where " &
      IF(B2="", "1=1", "Col2 = '"&B2&"'") & " and " &
      IF(D2="", "1=1", "Col3 >= "&D2) & " and " &
      IF(E2="", "1=1", "Col3 <= "&E2) & " and " &
      IF(G2="", "1=1", "Col4 = '"&G2&"'") & " and " &
      IF(I2="", "1=1", "Col10 contains '"&I2&"'"),
      1
    )
  )`;
  sh.getRange('A5').setFormula(formula);

  sh.setColumnWidths(1, 10, 160);
  sh.autoResizeColumns(1, 10);
  sh.getRange('H:H').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

function rebuildGalleryViewWithFilter() {
  const ss = SpreadsheetApp.getActive();
  const view = ss.getSheetByName(PCX.SHEETS.GALLERY_VIEW) || ss.insertSheet(PCX.SHEETS.GALLERY_VIEW);
  view.clear();

  view.getRange('A1:J1').setValues([[
    'Stamp ID','Country','Year','Denomination','Theme','Issue','Condition','Thumbnail','Open Image','Caption'
  ]]).setFontWeight('bold');
  view.setFrozenRows(1);

  const galName = PCX.SHEETS.GALLERY;
  const formula = `=ARRAYFORMULA({
    'Stamp Logging Assistant'!A:V,
    IFERROR(VLOOKUP('Stamp Logging Assistant'!A:A, ${galName}!A:E, 4, FALSE), "")
  })`;
  view.getRange('A2').setFormula(formula);

  if (view.getFilter()) view.getFilter().remove();
  view.getRange(1,1,1,10).createFilter();

  view.setColumnWidths(1, 10, 160);
  view.getRange('H:H').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

function setPickerDefaults_(){
  // Writes the constants into Script Properties so the sidebar can read them.
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PCX.PROP.PICKER_API_KEY, PC.PROP.PICKER_API_KEY);
  props.setProperty(PCX.PROP.PICKER_CLIENT_ID, PC.PROP.PICKER_CLIENT_ID);
  SpreadsheetApp.getUi().alert('Picker API key and Client ID saved to Script Properties.');
}

/***** Configure API Keys (Vision) *****/
function openConfigSidebar() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty(PC.PROP.VISION_API_KEY) || PC_DEFAULTS.VISION_API_KEY || '';
  const html = HtmlService.createHtmlOutput(`
    <div style="font: 13px/1.4 Roboto,Arial; padding:12px 14px; width: 340px;">
      <h3 style="margin:0 0 8px;">Philatelic Companion – API Keys</h3>
      <label>Google Cloud Vision API Key</label>
      <input id="vision" style="width:100%; margin:6px 0 12px;" value="${apiKey}" />
      <div style="display:flex; gap:8px;">
        <button onclick="save()" style="padding:6px 10px;">Save</button>
        <button onclick="google.script.host.close()" style="padding:6px 10px;">Close</button>
      </div>
      <p style="color:#666; margin-top:10px;">Enable Vision API on your GCP project, create an API key, and paste it here.</p>
      <script>
        function save(){
          const key = document.getElementById('vision').value.trim();
          google.script.run.withSuccessHandler(() => {
            google.script.host.close();
          }).setApiKeysFromSidebar_({ vision: key });
        }
      </script>
    </div>
  `).setTitle('Configure API Keys');
  SpreadsheetApp.getUi().showSidebar(html);
}

function setApiKeysFromSidebar_(form) {
  const props = PropertiesService.getScriptProperties();
  if (form && typeof form.vision === 'string') {
    props.setProperty(PC.PROP.VISION_API_KEY, form.vision.trim());
  }
}

/***** eBay export helpers *****/
function exportSelectionToEbay() {
  const ss   = SpreadsheetApp.getActive();
  const asst = ss.getSheetByName(PCX.SHEETS.ASST);
  const ebay = ss.getSheetByName(PCX.SHEETS.EBAY);
  if (!asst || !ebay) {
    SpreadsheetApp.getUi().alert('Assistant or eBay sheet not found.');
    return;
  }

  const sel = asst.getActiveRange();
  if (!sel || sel.getNumRows() < 1) {
    SpreadsheetApp.getUi().alert('Select one or more rows in the Assistant to export.');
    return;
  }

  const rows = sel.getValues();
  const out  = rows.map(r => {
    const stampId = r[0];
    const country = r[1], year = r[2], denom = r[3], cond = r[4];
    const catSys  = r[6], catNo = r[7], issue = r[8];
    const imgUrl  = findImageUrlById_(stampId) || '';

    const titleParts = [country, year, issue, catSys, catNo, denom, cond].filter(Boolean);
    const title = titleParts.join(' ').replace(/\s+/g, ' ').trim();

    const categoryId = getEbayCategoryId(title) || ''; // API lookup

    return [
      'Add',                // Action
      stampId,              // Custom label (SKU)
      categoryId,           // Category ID
      title,                // Title
      '',                   // UPC (blank if N/A)
      '',                   // Price
      1,                    // Quantity
      imgUrl,               // Item photo URL
      '',                   // Condition ID
      '',                   // Description
      'FixedPrice'          // Format
    ];
  });

  // Write after header rows
  let writeRow = Math.max(6, ebay.getLastRow() + 1);
  ebay.getRange(writeRow, 1, out.length, out[0].length).setValues(out);
  SpreadsheetApp.getUi().alert(`Exported ${out.length} row(s) to eBay sheet starting at row ${writeRow}.`);
}

function getEbayCategoryId(keywords) {
  const appId = 'joshuach-Recovere-SBX-b37694639-c7de5645'; // Replace with your eBay App ID
  const endpoint = 'https://svcs.ebay.com/services/search/FindingService/v1';
  const params = {
    'OPERATION-NAME': 'findItemsByKeywords',
    'SERVICE-VERSION': '1.0.0',
    'SECURITY-APPNAME': appId,
    'RESPONSE-DATA-FORMAT': 'JSON',
    'REST-PAYLOAD': '',
    'keywords': keywords,
    'paginationInput.entriesPerPage': 1
  };

  const query = Object.keys(params)
    .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k]))
    .join('&');

  const url = endpoint + '?' + query;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const data = JSON.parse(response.getContentText());

  try {
    const categoryId = data.findItemsByKeywordsResponse[0]
      .searchResult[0].item[0].primaryCategory[0].categoryId[0];
    return categoryId;
  } catch (e) {
    Logger.log('Category lookup failed for: ' + keywords);
    return '';
  }
}

/***** Photos readiness check (optional helper) *****/
function verifyPhotosImportSetup() {
  const requiredScope = 'https://www.googleapis.com/auth/photoslibrary.readonly';
  try {
    UrlFetchApp.fetch('https://photoslibrary.googleapis.com/v1/albums?pageSize=1');
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Google Photos import may not be authorized.\n\n' +
      'Go to Project Settings → Scopes and add:\n' + requiredScope +
      '\n\nThen re-authorize the script.'
    );
  }
}
