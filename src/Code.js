var CONFIG = {
  SHEET_NAME: '2026-27 Event Calendar',
  SCHOOL_YEAR_CELL_A1: 'A1',
  LOCATION_CELL_A1: 'G1',
  HEADER_SCAN_ROWS: 10,
  DAY_LABELS: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  TIMEZONE: 'America/New_York',
  SUNSET_FORMAT: 'HH:mm',
  INVALID_WEEK_TEXT: 'Invalid Week',
  NO_LOCATION_TEXT: 'No location set',
  INVALID_LOCATION_TEXT: 'Invalid Location'
};

/**
 * Adds menu actions to set up triggers and run backfill.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Troop Calendar')
    .addItem('Setup edit trigger', 'setupEventCalendarTrigger')
    .addItem('Set Location (Easy button)', 'openLocationDialog')
    .addItem('Backfill Day and Sunset', 'backfillAllRows')
    .addToUi();
}

/**
 * Opens modal dialog that can request browser geolocation on user click.
 */
function openLocationDialog() {
  var html = HtmlService.createHtmlOutputFromFile('LocationDialog')
    .setWidth(460)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Set Shared Location');
}

/**
 * Returns current shared location text from the configured location cell.
 *
 * @returns {{locationCell:string, currentValue:string}} Current sheet location metadata.
 */
function getLocationDialogContext() {
  var sheet = getTargetSheet_();
  if (!sheet) {
    throw new Error('Target sheet not found: ' + CONFIG.SHEET_NAME);
  }
  return {
    locationCell: CONFIG.LOCATION_CELL_A1,
    currentValue: String(sheet.getRange(CONFIG.LOCATION_CELL_A1).getDisplayValue() || '').trim()
  };
}

/**
 * Resolves browser coordinates into user-friendly location choices.
 *
 * @param {number} lat Latitude from browser geolocation.
 * @param {number} lng Longitude from browser geolocation.
 * @returns {{ok:boolean, cityState:string, zip:string, summary:string, error:string}}
 */
function resolveLocationChoicesFromCoordinates(lat, lng) {
  var coords = validateCoordinates_(lat, lng);
  var geocoder = Maps.newGeocoder();
  var result = geocoder.reverseGeocode(coords.lat, coords.lng);
  if (!result || result.status !== 'OK' || !result.results || !result.results.length) {
    return {
      ok: false,
      cityState: '',
      zip: '',
      summary: '',
      error: 'Could not resolve your location. You can still type City, ST or ZIP manually.'
    };
  }

  var components = result.results[0].address_components || [];
  var city = extractAddressComponent_(components, 'locality') || extractAddressComponent_(components, 'postal_town');
  var state = extractAddressComponent_(components, 'administrative_area_level_1');
  var zip = extractAddressComponent_(components, 'postal_code');
  var cityState = formatCityState_(city, state);
  var summaryParts = [];

  if (cityState) {
    summaryParts.push(cityState);
  }
  if (zip) {
    summaryParts.push(zip);
  }

  if (!cityState && !zip) {
    return {
      ok: false,
      cityState: '',
      zip: '',
      summary: '',
      error: 'Location found, but City/State or ZIP were not available. Please type a value manually.'
    };
  }

  return {
    ok: true,
    cityState: cityState,
    zip: zip,
    summary: summaryParts.join(' | '),
    error: ''
  };
}

/**
 * Saves chosen shared location text into the configured location cell and recalculates sunset values.
 *
 * @param {string} locationText Chosen location text (City, ST or ZIP).
 * @returns {{ok:boolean, locationCell:string, value:string, recalculated:boolean, warning:string}} Result metadata.
 */
function setSharedLocationValue(locationText) {
  var value = String(locationText || '').trim();
  if (!value) {
    throw new Error('Location value cannot be empty.');
  }

  var sheet = getTargetSheet_();
  if (!sheet) {
    throw new Error('Target sheet not found: ' + CONFIG.SHEET_NAME);
  }

  sheet.getRange(CONFIG.LOCATION_CELL_A1).setValue(value);

  var headerInfo = getHeaderInfo_(sheet);
  var recalculated = false;
  var warning = '';
  if (headerInfo) {
    backfillRowsWithHeaderInfo_(sheet, headerInfo);
    recalculated = true;
  } else {
    warning = 'Location saved, but sunset recalculation was skipped because headers Week, Day, Sunset were not found on the target sheet.';
  }

  return {
    ok: true,
    locationCell: CONFIG.LOCATION_CELL_A1,
    value: value,
    recalculated: recalculated,
    warning: warning
  };
}

/**
 * Creates an installable on-edit trigger for this spreadsheet.
 * Run once after script deployment.
 */
function setupEventCalendarTrigger() {
  var spreadsheetId = SpreadsheetApp.getActive().getId();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    if (t.getHandlerFunction() === 'installedOnEdit') {
      return;
    }
  }
  ScriptApp.newTrigger('installedOnEdit')
    .forSpreadsheet(spreadsheetId)
    .onEdit()
    .create();
}

/**
 * Installable edit trigger entrypoint.
 * Updates Day and Sunset based on Week text and school-year anchor.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Edit event.
 */
function installedOnEdit(e) {
  if (!e || !e.range) {
    return;
  }

  var sheet = e.range.getSheet();
  if (!isTargetSheet_(sheet)) {
    return;
  }

  var headerInfo = getHeaderInfo_(sheet);
  if (!headerInfo) {
    return;
  }

  var row = e.range.getRow();
  var col = e.range.getColumn();
  var schoolYearCell = a1ToRowCol_(CONFIG.SCHOOL_YEAR_CELL_A1);
  var locationCell = a1ToRowCol_(CONFIG.LOCATION_CELL_A1);

  if (row === schoolYearCell.row && col === schoolYearCell.col) {
    backfillAllRows();
    return;
  }

  if (row === locationCell.row && col === locationCell.col) {
    backfillAllRows();
    return;
  }

  if (col !== headerInfo.weekCol || row <= headerInfo.headerRow) {
    return;
  }

  updateRowFromWeek_(sheet, row, headerInfo);
}

/**
 * Backfills all rows under the header using current Week values.
 */
function backfillAllRows() {
  var sheet = getTargetSheet_();
  if (!sheet) {
    throw new Error('Target sheet not found: ' + CONFIG.SHEET_NAME);
  }

  var headerInfo = getHeaderInfo_(sheet);
  if (!headerInfo) {
    throw new Error('Could not locate required headers: Week, Day, Sunset.');
  }

  backfillRowsWithHeaderInfo_(sheet, headerInfo);
}

/**
 * Backfills rows for a sheet once header metadata has already been resolved.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet instance.
 * @param {{headerRow:number, weekCol:number, dayCol:number, sunsetCol:number}} headerInfo Header metadata.
 */
function backfillRowsWithHeaderInfo_(sheet, headerInfo) {
  if (!sheet || !headerInfo) {
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= headerInfo.headerRow) {
    return;
  }

  for (var row = headerInfo.headerRow + 1; row <= lastRow; row++) {
    updateRowFromWeek_(sheet, row, headerInfo);
  }
}

/**
 * Updates Day and Sunset for one row using Week text.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet instance.
 * @param {number} row Row number.
 * @param {{headerRow:number, weekCol:number, dayCol:number, sunsetCol:number}} headerInfo Header metadata.
 */
function updateRowFromWeek_(sheet, row, headerInfo) {
  var weekValue = sheet.getRange(row, headerInfo.weekCol).getValue();
  if (weekValue === '' || weekValue === null) {
    writeIfChanged_(sheet, row, headerInfo.dayCol, '');
    writeIfChanged_(sheet, row, headerInfo.sunsetCol, '');
    return;
  }

  var schoolYearText = sheet.getRange(CONFIG.SCHOOL_YEAR_CELL_A1).getDisplayValue();
  var parsed = parseWeekValue_(weekValue, schoolYearText);

  if (!parsed.ok) {
    writeIfChanged_(sheet, row, headerInfo.dayCol, CONFIG.INVALID_WEEK_TEXT);
    writeIfChanged_(sheet, row, headerInfo.sunsetCol, CONFIG.INVALID_WEEK_TEXT);
    return;
  }

  var dayText = buildDayText_(parsed.startDate, parsed.endDate);
  var sunsetText = getCachedSunsetForDate_(parsed.startDate, sheet);

  writeIfChanged_(sheet, row, headerInfo.dayCol, dayText);
  writeIfChanged_(sheet, row, headerInfo.sunsetCol, sunsetText);
}

/**
 * Parses Week values into resolved start/end Date objects using A1 school year.
 * Supported examples: "Aug 9", "Aug 10-12", "Aug 30-Sep 2".
 *
 * @param {*} weekValue Week cell value.
 * @param {string} schoolYearText A1 value like "2026-27".
 * @returns {{ok:boolean, startDate:Date, endDate:Date, error:string}}
 */
function parseWeekValue_(weekValue, schoolYearText) {
  if (weekValue instanceof Date && !isNaN(weekValue.getTime())) {
    return { ok: true, startDate: stripTime_(weekValue), endDate: stripTime_(weekValue), error: '' };
  }

  var text = String(weekValue || '').trim();
  if (!text) {
    return { ok: false, startDate: null, endDate: null, error: 'Empty week value' };
  }

  text = text.replace(/[\u2012\u2013\u2014\u2015]/g, '-').replace(/\s+/g, ' ');

  var pattern = /^([A-Za-z]+)\s+(\d{1,2})(?:\s*-\s*([A-Za-z]+)?\s*(\d{1,2}))?$/;
  var match = text.match(pattern);
  if (!match) {
    return { ok: false, startDate: null, endDate: null, error: 'Unrecognized format' };
  }

  var startMonth = parseMonthToken_(match[1]);
  var startDay = Number(match[2]);
  if (startMonth < 0 || !isValidDayOfMonth_(startDay)) {
    return { ok: false, startDate: null, endDate: null, error: 'Invalid start month/day' };
  }

  var startYear = resolveSchoolYearForMonth_(schoolYearText, startMonth);
  var startDate = new Date(startYear, startMonth, startDay);
  if (!isExactDate_(startDate, startYear, startMonth, startDay)) {
    return { ok: false, startDate: null, endDate: null, error: 'Invalid start date' };
  }

  if (!match[4]) {
    return { ok: true, startDate: startDate, endDate: startDate, error: '' };
  }

  var endMonth = match[3] ? parseMonthToken_(match[3]) : startMonth;
  var endDay = Number(match[4]);
  if (endMonth < 0 || !isValidDayOfMonth_(endDay)) {
    return { ok: false, startDate: null, endDate: null, error: 'Invalid end month/day' };
  }

  var endYear = resolveSchoolYearForMonth_(schoolYearText, endMonth);
  var endDate = new Date(endYear, endMonth, endDay);
  if (!isExactDate_(endDate, endYear, endMonth, endDay)) {
    return { ok: false, startDate: null, endDate: null, error: 'Invalid end date' };
  }

  if (endDate.getTime() < startDate.getTime()) {
    return { ok: false, startDate: null, endDate: null, error: 'End before start' };
  }

  return { ok: true, startDate: startDate, endDate: endDate, error: '' };
}

/**
 * Converts parsed dates to Day output text.
 *
 * @param {Date} startDate Resolved start date.
 * @param {Date} endDate Resolved end date.
 * @returns {string} Day text like "Fri" or "Fri-Sun".
 */
function buildDayText_(startDate, endDate) {
  var startLabel = CONFIG.DAY_LABELS[startDate.getDay()];
  var endLabel = CONFIG.DAY_LABELS[endDate.getDay()];
  return startDate.getTime() === endDate.getTime() ? startLabel : startLabel + '-' + endLabel;
}

/**
 * Gets sunset value from cache or API, formatted for configured timezone.
 *
 * @param {Date} date Date to query.
 * @returns {string} Sunset time in HH:mm.
 */
function getCachedSunsetForDate_(date, sheet) {
  var locationInfo = resolveSunsetLocation_(sheet);
  if (!locationInfo.ok) {
    if (locationInfo.reason === 'missing') {
      return CONFIG.NO_LOCATION_TEXT;
    }
    return CONFIG.INVALID_LOCATION_TEXT;
  }

  var dateKey = toDateKey_(date);
  var propKey = 'sunset:' + locationInfo.cacheKey + ':' + dateKey;
  var props = PropertiesService.getDocumentProperties();
  var cached = props.getProperty(propKey);
  if (cached) {
    return cached;
  }

  var url = 'https://api.sunrise-sunset.org/json?lat=' +
    encodeURIComponent(locationInfo.lat) +
    '&lng=' + encodeURIComponent(locationInfo.lng) +
    '&date=' + encodeURIComponent(dateKey);
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    return CONFIG.INVALID_WEEK_TEXT;
  }

  var data = JSON.parse(response.getContentText());
  if (!data || data.status !== 'OK' || !data.results || !data.results.sunset) {
    return CONFIG.INVALID_WEEK_TEXT;
  }

  var utcString = dateKey + ' ' + data.results.sunset + ' UTC';
  var dateObj = new Date(utcString);
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
    return CONFIG.INVALID_WEEK_TEXT;
  }

  var formatted = Utilities.formatDate(dateObj, CONFIG.TIMEZONE, CONFIG.SUNSET_FORMAT);
  props.setProperty(propKey, formatted);
  return formatted;
}

/**
 * Resolves sunset lookup location from configured cell or geocoded input.
 * Supports US ZIP (5-digit or ZIP+4) and city/state text.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet instance.
 * @returns {{ok:boolean, lat:number, lng:number, cacheKey:string, error:string, reason:string}}
 */
function resolveSunsetLocation_(sheet) {
  var locationText = String(sheet.getRange(CONFIG.LOCATION_CELL_A1).getDisplayValue() || '').trim();
  if (!locationText) {
    return {
      ok: false,
      lat: 0,
      lng: 0,
      cacheKey: 'none',
      error: 'No location set',
      reason: 'missing'
    };
  }

  var normalized = normalizeLocationInput_(locationText);
  var props = PropertiesService.getDocumentProperties();
  var geoKey = 'geocode:' + normalized;
  var cachedGeo = props.getProperty(geoKey);
  if (cachedGeo) {
    try {
      var parsedCached = JSON.parse(cachedGeo);
      if (
        parsedCached &&
        Number.isFinite(parsedCached.lat) &&
        Number.isFinite(parsedCached.lng)
      ) {
        return {
          ok: true,
          lat: parsedCached.lat,
          lng: parsedCached.lng,
          cacheKey: normalized,
          error: '',
          reason: ''
        };
      }
    } catch (err) {
      // Ignore cache parse errors and continue with a fresh geocode lookup.
    }
  }

  var geocoder = Maps.newGeocoder();
  var result = geocoder.geocode(locationText);
  if (!result || result.status !== 'OK' || !result.results || !result.results.length) {
    return {
      ok: false,
      lat: 0,
      lng: 0,
      cacheKey: normalized,
      error: 'Geocode lookup failed',
      reason: 'invalid'
    };
  }

  var location = result.results[0] && result.results[0].geometry && result.results[0].geometry.location;
  if (!location || !Number.isFinite(location.lat) || !Number.isFinite(location.lng)) {
    return {
      ok: false,
      lat: 0,
      lng: 0,
      cacheKey: normalized,
      error: 'Geocode did not return coordinates',
      reason: 'invalid'
    };
  }

  props.setProperty(geoKey, JSON.stringify({ lat: location.lat, lng: location.lng }));
  return {
    ok: true,
    lat: location.lat,
    lng: location.lng,
    cacheKey: normalized,
    error: '',
    reason: ''
  };
}

/**
 * Normalizes location input text for stable property keys.
 *
 * @param {string} value Raw location input.
 * @returns {string} Normalized key-safe text.
 */
function normalizeLocationInput_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^a-z0-9,\- ]/g, '')
    .replace(/[ ,\-]+/g, '_');
}

/**
 * Validates lat/lng values from browser geolocation payload.
 *
 * @param {*} lat Latitude candidate.
 * @param {*} lng Longitude candidate.
 * @returns {{lat:number, lng:number}} Normalized coordinates.
 */
function validateCoordinates_(lat, lng) {
  var latNum = Number(lat);
  var lngNum = Number(lng);
  var valid = Number.isFinite(latNum) && Number.isFinite(lngNum) && latNum >= -90 && latNum <= 90 && lngNum >= -180 && lngNum <= 180;
  if (!valid) {
    throw new Error('Invalid coordinates received from browser geolocation.');
  }
  return { lat: latNum, lng: lngNum };
}

/**
 * Extracts short_name for a given Google address component type.
 *
 * @param {Array<{types:string[],short_name:string}>} components Geocoder components.
 * @param {string} targetType Component type to find.
 * @returns {string} Component short name or empty string.
 */
function extractAddressComponent_(components, targetType) {
  var list = Array.isArray(components) ? components : [];
  for (var i = 0; i < list.length; i++) {
    var entry = list[i];
    var types = entry && entry.types;
    if (Array.isArray(types) && types.indexOf(targetType) >= 0) {
      return String(entry.short_name || '').trim();
    }
  }
  return '';
}

/**
 * Formats city/state text.
 *
 * @param {string} city Locality value.
 * @param {string} state State code/value.
 * @returns {string} Formatted city/state or empty string.
 */
function formatCityState_(city, state) {
  var cityPart = String(city || '').trim();
  var statePart = String(state || '').trim();
  if (cityPart && statePart) {
    return cityPart + ', ' + statePart;
  }
  if (cityPart) {
    return cityPart;
  }
  return '';
}

/**
 * Resolves year based on school-year text where Aug-Dec use start year and Jan-Jul use start year + 1.
 *
 * @param {string} schoolYearText A1 value like "2026-27".
 * @param {number} monthIndex Zero-based month index.
 * @returns {number} Resolved year.
 */
function resolveSchoolYearForMonth_(schoolYearText, monthIndex) {
  var startYear = parseSchoolYearStart_(schoolYearText);
  return monthIndex >= 7 ? startYear : startYear + 1;
}

/**
 * Parses start year from school-year text.
 * Accepts "2026-27" and "2026".
 *
 * @param {string} schoolYearText A1 school-year text.
 * @returns {number} Start year.
 */
function parseSchoolYearStart_(schoolYearText) {
  var text = String(schoolYearText || '').trim();
  var match = text.match(/(\d{4})/);
  if (!match) {
    throw new Error('A1 must contain school year text like 2026-27.');
  }
  return Number(match[1]);
}

/**
 * Parses month token to zero-based month index.
 *
 * @param {string} token Month text.
 * @returns {number} Month index or -1 if unknown.
 */
function parseMonthToken_(token) {
  var value = String(token || '').toLowerCase();
  var map = {
    jan: 0, january: 0,
    feb: 1, february: 1,
    mar: 2, march: 2,
    apr: 3, april: 3,
    may: 4,
    jun: 5, june: 5,
    jul: 6, july: 6,
    aug: 7, august: 7,
    sep: 8, sept: 8, september: 8,
    oct: 9, october: 9,
    nov: 10, november: 10,
    dec: 11, december: 11
  };
  return Object.prototype.hasOwnProperty.call(map, value) ? map[value] : -1;
}

/**
 * Finds header row and key columns by header names.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet instance.
 * @returns {{headerRow:number, weekCol:number, dayCol:number, sunsetCol:number}|null}
 */
function getHeaderInfo_(sheet) {
  var maxRows = Math.min(CONFIG.HEADER_SCAN_ROWS, sheet.getLastRow());
  if (maxRows < 1) {
    return null;
  }

  var maxCols = Math.max(1, sheet.getLastColumn());
  var values = sheet.getRange(1, 1, maxRows, maxCols).getDisplayValues();

  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    var weekCol = findHeaderIndex_(row, 'Week');
    var dayCol = findHeaderIndex_(row, 'Day');
    var sunsetCol = findHeaderIndex_(row, 'Sunset');
    if (weekCol > 0 && dayCol > 0 && sunsetCol > 0) {
      return {
        headerRow: r + 1,
        weekCol: weekCol,
        dayCol: dayCol,
        sunsetCol: sunsetCol
      };
    }
  }

  return null;
}

/**
 * Returns 1-based index of target header in row values.
 *
 * @param {string[]} rowValues Header row values.
 * @param {string} target Header target.
 * @returns {number} Column number or 0 if not found.
 */
function findHeaderIndex_(rowValues, target) {
  var normTarget = String(target || '').trim().toLowerCase();
  for (var c = 0; c < rowValues.length; c++) {
    if (String(rowValues[c] || '').trim().toLowerCase() === normTarget) {
      return c + 1;
    }
  }
  return 0;
}

/**
 * Writes value to cell only when changed.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet instance.
 * @param {number} row Row number.
 * @param {number} col Column number.
 * @param {string} value Desired value.
 */
function writeIfChanged_(sheet, row, col, value) {
  var range = sheet.getRange(row, col);
  var current = range.getDisplayValue();
  var next = String(value == null ? '' : value);
  if (current !== next) {
    range.setValue(next);
  }
}

/**
 * Returns target sheet by config.
 * Falls back to active sheet when configured name is not present.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getTargetSheet_() {
  var ss = SpreadsheetApp.getActive();
  var named = ss.getSheetByName(CONFIG.SHEET_NAME);
  return named || ss.getActiveSheet();
}

/**
 * Checks whether edited sheet is the configured target.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet instance.
 * @returns {boolean} True when sheet is target.
 */
function isTargetSheet_(sheet) {
  var target = getTargetSheet_();
  return !!target && target.getSheetId() === sheet.getSheetId();
}

/**
 * Parses A1 notation containing one cell to row/col object.
 *
 * @param {string} a1 Cell reference.
 * @returns {{row:number,col:number}}
 */
function a1ToRowCol_(a1) {
  var match = String(a1 || '').match(/^([A-Za-z]+)(\d+)$/);
  if (!match) {
    throw new Error('Invalid A1 notation: ' + a1);
  }
  var letters = match[1].toUpperCase();
  var row = Number(match[2]);
  var col = 0;
  for (var i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { row: row, col: col };
}

/**
 * Formats Date as YYYY-MM-DD.
 *
 * @param {Date} date Date object.
 * @returns {string} Date key.
 */
function toDateKey_(date) {
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
}

/**
 * Removes time component from Date.
 *
 * @param {Date} date Input date.
 * @returns {Date} Date at local midnight.
 */
function stripTime_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * Checks day number bounds.
 *
 * @param {number} day Day number.
 * @returns {boolean} True when valid range.
 */
function isValidDayOfMonth_(day) {
  return Number.isFinite(day) && day >= 1 && day <= 31;
}

/**
 * Confirms Date constructor did not overflow month/day.
 *
 * @param {Date} date Constructed date.
 * @param {number} year Expected year.
 * @param {number} month Expected month index.
 * @param {number} day Expected day.
 * @returns {boolean} True when exact.
 */
function isExactDate_(date, year, month, day) {
  return date.getFullYear() === year && date.getMonth() === month && date.getDate() === day;
}
