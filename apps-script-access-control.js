const ACCESS_SPREADSHEET_ID = '1QLEt5YP14t5H3lXfwcjjXzOuZDx4NNeBEOPzlENKr9E';
const ACCESS_SHEET_NAME = 'alergenos';
const SESSION_SHEET_NAME = 'sesiones';
const SESSION_TTL_MINUTES = 60;
const SUGGESTION_RECIPIENTS = 'aina@sagardi.com,jlungidos@sagardi.com';
const SUGGESTION_COOLDOWN_SECONDS = 30;

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData && e.postData.contents || '{}');
    return jsonResponse(handlePayload(payload));
  } catch (error) {
    return jsonResponse({ ok: false, error: 'SERVER_ERROR', message: String(error && error.message || error) });
  }
}

function doGet(e) {
  try {
    const payload = JSON.parse(e.parameter && e.parameter.payload || '{}');
    const callback = String(e.parameter && e.parameter.callback || '').replace(/[^\w.$]/g, '');
    const result = handlePayload(payload);
    if (callback) return javascriptResponse(`${callback}(${JSON.stringify(result)});`);
    return jsonResponse(result);
  } catch (error) {
    const result = { ok: false, error: 'SERVER_ERROR', message: String(error && error.message || error) };
    const callback = String(e.parameter && e.parameter.callback || '').replace(/[^\w.$]/g, '');
    if (callback) return javascriptResponse(`${callback}(${JSON.stringify(result)});`);
    return jsonResponse(result);
  }
}

function handlePayload(payload) {
  if (payload.type === 'access-profile-update') return updateAccessProfile(payload.profile);
  if (payload.type === 'session-start') return startSession(payload.session);
  if (payload.type === 'session-end') return endSession(payload.session);
  if (payload.type === 'suggestion') return sendSuggestion(payload);
  return { ok: false, error: 'UNKNOWN_ACTION' };
}

function sendSuggestion(payload) {
  const suggestion = payload && payload.suggestion || {};
  const local = cleanSuggestionText(suggestion.local, 120);
  const name = cleanSuggestionText(suggestion.name, 120);
  const observations = cleanSuggestionText(suggestion.observations, 1500, true);
  if (!local || !name || !observations) return { ok: false, error: 'MISSING_FIELDS' };

  const sessionToken = cleanSuggestionText(payload.sessionToken, 180) || `${local}|${name}`;
  const cache = CacheService.getScriptCache();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sessionToken);
  const rateKey = `suggestion-${Utilities.base64EncodeWebSafe(digest).slice(0, 40)}`;
  if (cache.get(rateKey)) return { ok: false, error: 'RATE_LIMIT' };
  if (MailApp.getRemainingDailyQuota() < 2) return { ok: false, error: 'MAIL_QUOTA' };

  const appUser = cleanSuggestionText(suggestion.appUser, 120) || 'No indicado';
  const sentAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/Madrid', 'dd/MM/yyyy HH:mm');
  const subject = `Sugerencia App Alérgenos - ${local}`;
  const plainBody = [
    'Nueva sugerencia recibida desde la App de Alérgenos',
    '',
    `Local: ${local}`,
    `Persona: ${name}`,
    `Usuario de la app: ${appUser}`,
    `Fecha: ${sentAt}`,
    '',
    'Sugerencia / observaciones:',
    observations
  ].join('\n');
  const htmlBody = [
    '<h2>Nueva sugerencia de la App de Alérgenos</h2>',
    `<p><strong>Local:</strong> ${escapeSuggestionHtml(local)}<br>`,
    `<strong>Persona:</strong> ${escapeSuggestionHtml(name)}<br>`,
    `<strong>Usuario de la app:</strong> ${escapeSuggestionHtml(appUser)}<br>`,
    `<strong>Fecha:</strong> ${escapeSuggestionHtml(sentAt)}</p>`,
    '<h3>Sugerencia / observaciones</h3>',
    `<p style="white-space:pre-wrap">${escapeSuggestionHtml(observations)}</p>`
  ].join('');

  MailApp.sendEmail({
    to: SUGGESTION_RECIPIENTS,
    subject,
    body: plainBody,
    htmlBody,
    name: 'App de Alérgenos Sagardi'
  });
  cache.put(rateKey, '1', SUGGESTION_COOLDOWN_SECONDS);
  return { ok: true };
}

function cleanSuggestionText(value, maxLength, preserveLines) {
  const text = String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
    .trim();
  const cleaned = preserveLines
    ? text.split('\n').map(line => line.replace(/[ \t]+/g, ' ').trim()).join('\n')
    : text.replace(/\s+/g, ' ');
  return cleaned.slice(0, maxLength);
}

function escapeSuggestionHtml(value) {
  return String(value || '').replace(/[&<>"']/g, character => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[character]);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function javascriptResponse(source) {
  return ContentService
    .createTextOutput(source)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function accessSheet() {
  return SpreadsheetApp.openById(ACCESS_SPREADSHEET_ID).getSheetByName(ACCESS_SHEET_NAME);
}

function sessionSheet() {
  const ss = SpreadsheetApp.openById(ACCESS_SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SESSION_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SESSION_SHEET_NAME);
    sheet.appendRow(['profileId', 'user', 'slot', 'token', 'startedAt', 'lastSeenAt']);
  }
  return sheet;
}

function headerMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const counts = {};
  const map = {};
  headers.forEach((header, index) => {
    const key = normalizeKey(header);
    if (!key) return;
    const count = counts[key] || 0;
    counts[key] = count + 1;
    if (count === 0) map[key] = index + 1;
    if (count > 0) map[`${key}.${count}`] = index + 1;
  });
  return map;
}

function normalizeKey(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function column(map, names) {
  for (const name of names) {
    const key = normalizeKey(name);
    if (map[key]) return map[key];
  }
  return 0;
}

function updateAccessProfile(profile) {
  if (!profile || typeof profile.rowIndex !== 'number') return { ok: false, error: 'BAD_PROFILE' };
  const sheet = accessSheet();
  const map = headerMap(sheet);
  const row = profile.rowIndex + 1;
  setIfColumn(sheet, row, column(map, ['TIPO']), profile.tipo);
  setIfColumn(sheet, row, column(map, ['Nº', 'NO', 'NUM', 'N']), profile.number);
  setIfColumn(sheet, row, column(map, ['LOCAL']), profile.local);
  setIfColumn(sheet, row, column(map, ['VINCULADO A', 'VINCULADO']), profile.linked);
  (profile.credentials || []).forEach(credential => {
    const suffix = credential.slot === 0 ? '' : `.${credential.slot}`;
    setIfColumn(sheet, row, column(map, [`Usuario${suffix}`]), credential.user);
    setIfColumn(sheet, row, column(map, [`Contra${suffix}`]), credential.pass);
    setIfColumn(sheet, row, column(map, [`Act${suffix}`]), credential.active ? true : false);
    setIfColumn(sheet, row, column(map, [`Personas${suffix}`]), credential.people || '');
  });
  return { ok: true };
}

function setIfColumn(sheet, row, col, value) {
  if (col > 0) sheet.getRange(row, col).setValue(value);
}

function startSession(session) {
  if (!session || !session.profileId || !session.user || !session.sessionToken) return { ok: false, error: 'BAD_SESSION' };
  const sheet = sessionSheet();
  expireOldSessions(sheet);
  const limit = Math.max(1, Number(session.peopleLimit || 1));
  const rows = sheet.getDataRange().getValues();
  const active = rows.slice(1).filter(row =>
    String(row[0]) === String(session.profileId) &&
    normalizeKey(row[1]) === normalizeKey(session.user) &&
    String(row[2]) === String(session.credentialSlot)
  );
  const sameToken = active.some(row => String(row[3]) === String(session.sessionToken));
  if (!sameToken && active.length >= limit) return { ok: false, error: 'SESSION_LIMIT', limit };
  const now = new Date();
  if (sameToken) {
    const index = rows.findIndex(row => String(row[3]) === String(session.sessionToken));
    if (index > 0) sheet.getRange(index + 1, 6).setValue(now);
  } else {
    sheet.appendRow([session.profileId, session.user, session.credentialSlot, session.sessionToken, now, now]);
  }
  return { ok: true };
}

function endSession(session) {
  if (!session || !session.sessionToken) return { ok: true };
  const sheet = sessionSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i -= 1) {
    if (String(rows[i][3]) === String(session.sessionToken)) sheet.deleteRow(i + 1);
  }
  return { ok: true };
}

function expireOldSessions(sheet) {
  const rows = sheet.getDataRange().getValues();
  const cutoff = Date.now() - SESSION_TTL_MINUTES * 60 * 1000;
  for (let i = rows.length - 1; i >= 1; i -= 1) {
    const lastSeen = rows[i][5] instanceof Date ? rows[i][5].getTime() : 0;
    if (!lastSeen || lastSeen < cutoff) sheet.deleteRow(i + 1);
  }
}
