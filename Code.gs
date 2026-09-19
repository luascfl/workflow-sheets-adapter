const CONFIG_PROPERTY = 'WORKFLOW_SHEETS_ADAPTER_CONFIG';

const CENTRAL_SPREADSHEET_ID = '1ftwnYkccSpqyY86qBzyXyzQnfog_mBhb2QqphqtL9JE';

const SHEET_SPECS = Object.freeze({
  alertasCiee: {
    label: 'CIEE',
    allowedRanges: [
      'AlertasEstagio!A:A',
      'AlertasEstagio!A:B',
      'AlertasEstagio!D:D',
      'AlertasAprendiz!A:A',
      'AlertasAprendiz!A:B',
      'AlertasAprendiz!D:D',
    ],
  },
  alertasEspro: { label: 'ESPRO TAQE', allowedRanges: ['Alertas!A:A', 'Alertas!A:B'] },
  alertasIsbet: { label: 'ISBET', allowedRanges: ['Alertas!A:A', 'Alertas!A:B'] },
  alertasIel: { label: 'IEL Bahia', allowedRanges: ['Alertas!A:A', 'Alertas!A:B'] },
  alertasStartCarreiras: { label: 'Start Carreiras', allowedRanges: ['Alertas!A:A', 'Alertas!A:B'] },
  alertasLinkedin: { label: 'LinkedIn', allowedRanges: ["Alertas!K3:K"] },
  alertasCentral: { label: 'Central de alertas', allowedRanges: [] },
});

function doGet(event) {
  const request = event && event.parameter ? event.parameter : {};
  if (request.ui === 'setup') {
    return HtmlService.createHtmlOutputFromFile('Setup').setTitle('Workflow Sheets Adapter');
  }
  return jsonResponse_(handleRequest_(request));
}

function doPost(event) {
  const request = event && event.postData && event.postData.contents
    ? JSON.parse(event.postData.contents)
    : {};
  return jsonResponse_(handleRequest_(request));
}

function getSetupState() {
  const config = recoverCentralOnlyState_();
  return {
    endpoint: ScriptApp.getService().getUrl(),
    accessToken: config.accessToken || '',
    sourcesDeleted: Boolean(config.lastMerge && config.lastMerge.sourcesDeletedAt),
    sheets: Object.keys(SHEET_SPECS).map((key) => ({
      key,
      label: SHEET_SPECS[key].label,
      isCentral: key === 'alertasCentral',
      configured: Boolean(config.sheets[key]),
      spreadsheetUrl: config.sheets[key]
        ? `https://docs.google.com/spreadsheets/d/${config.sheets[key].spreadsheetId}/edit`
        : '',
    })),
  };
}

function saveSetupState(form) {
  saveConfig_(buildConfig_(form, false));
  return getSetupState();
}

function rotateAccessToken() {
  const config = getConfig_();
  config.accessToken = createAccessToken_();
  saveConfig_(config);
  return getSetupState();
}

function createCentralWorkbook(form) {
  const config = buildConfig_(form, true);
  const existingId = config.sheets.alertasCentral && config.sheets.alertasCentral.spreadsheetId;
  if (!existingId) {
    const spreadsheet = SpreadsheetApp.create('$$alertasCentral');
    const instructions = spreadsheet.getActiveSheet();
    instructions.setName('Instruções');
    instructions.getRange('A1:B4').setValues([
      ['Central de alertas', 'Gerada pelo Workflow Sheets Adapter'],
      ['Política de duplicatas', 'Todos os registros foram preservados com plataforma e aba de origem.'],
      ['Fontes', 'CIEE, ESPRO TAQE, ISBET, IEL Bahia, Start Carreiras e LinkedIn.'],
      ['Uso', 'Não edite as abas importadas antes de validar a migração.'],
    ]);
    config.sheets.alertasCentral = { spreadsheetId: spreadsheet.getId() };
  }
  saveConfig_(config);
  return getSetupState();
}

function mergeConfiguredSheets() {
  const config = getConfig_();
  const missing = sourceKeys_().filter((key) => !config.sheets[key]);
  if (missing.length) {
    throw new Error(`Configure as fontes antes de mesclar: ${missing.map((key) => SHEET_SPECS[key].label).join(', ')}.`);
  }
  if (!config.sheets.alertasCentral) {
    throw new Error('Crie a planilha central antes de mesclar.');
  }

  const central = SpreadsheetApp.openById(config.sheets.alertasCentral.spreadsheetId);
  clearImportedTabs_(central);
  const result = [];

  sourceKeys_().forEach((sourceKey) => {
    const source = SpreadsheetApp.openById(config.sheets[sourceKey].spreadsheetId);
    source.getSheets().forEach((sourceTab) => {
      const values = sourceTab.getDataRange().getValues();
      if (values.length === 0 || (values.length === 1 && values[0].every((value) => value === ''))) {
        return;
      }
      const destinationName = createDestinationName_(central, SHEET_SPECS[sourceKey].label, sourceTab.getName());
      const destination = central.insertSheet(destinationName);
      const header = values[0];
      const rows = values.slice(1).map((row) => [SHEET_SPECS[sourceKey].label, sourceTab.getName(), ...row]);
      destination.getRange(1, 1, 1, header.length + 2)
        .setValues([['Origem', 'Aba de origem', ...header]]);
      if (rows.length) {
        destination.getRange(2, 1, rows.length, header.length + 2).setValues(rows);
      }
      destination.setFrozenRows(1);
      destination.autoResizeColumns(1, Math.min(header.length + 2, 12));
      result.push({ source: SHEET_SPECS[sourceKey].label, tab: sourceTab.getName(), rows: rows.length });
    });
  });

  const merge = {
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${central.getId()}/edit`,
    importedTabs: result,
    importedRows: result.reduce((total, item) => total + item.rows, 0),
  };
  config.lastMerge = { centralSpreadsheetId: central.getId(), sourceKeys: sourceKeys_(), timestamp: new Date().toISOString() };
  saveConfig_(config);
  return merge;
}

function permanentlyDeleteMergedSources() {
  const config = getConfig_();
  const merge = config.lastMerge;
  if (!merge || !config.sheets.alertasCentral || merge.centralSpreadsheetId !== config.sheets.alertasCentral.spreadsheetId) {
    throw new Error('Mescle novamente as fontes antes de excluí-las.');
  }

  const sourceKeys = sourceKeys_();
  const sourceIds = sourceKeys.map((key) => config.sheets[key].spreadsheetId);
  if (new Set(sourceIds).size !== sourceIds.length || sourceIds.includes(config.sheets.alertasCentral.spreadsheetId)) {
    throw new Error('A configuração das fontes é inválida para exclusão.');
  }

  sourceIds.forEach((spreadsheetId) => {
    // Forces Apps Script's Drive authorization flow before the irreversible REST call.
    DriveApp.getFileById(spreadsheetId);
    const response = UrlFetchApp.fetch(`https://www.googleapis.com/drive/v3/files/${spreadsheetId}`, {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      method: 'delete',
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() !== 204) {
      throw new Error(
        `Falha ao excluir a fonte ${spreadsheetId}: HTTP ${response.getResponseCode()}. ${response.getContentText()}`,
      );
    }
  });

  sourceKeys.forEach((key) => delete config.sheets[key]);
  config.lastMerge = { ...merge, sourcesDeletedAt: new Date().toISOString() };
  saveConfig_(config);
  return { deletedSheets: sourceKeys.map((key) => SHEET_SPECS[key].label) };
}

function inspectSourceDeletionAccess() {
  const config = getConfig_();
  return sourceKeys_().map((key) => {
    const spreadsheetId = config.sheets[key].spreadsheetId;
    const file = DriveApp.getFileById(spreadsheetId);
    return {
      label: SHEET_SPECS[key].label,
      spreadsheetId,
      owner: file.getOwner().getEmail(),
      trashed: file.isTrashed(),
      url: file.getUrl(),
    };
  });
}
function handleRequest_(request) {
  try {
    requireAccessToken_(request);
    switch (String(request.operation || 'status')) {
      case 'status':
        return { ok: true, configuredSheets: Object.keys(getConfig_().sheets) };
      case 'read':
        return { ok: true, values: getAuthorizedRange_(request).getValues() };
      case 'append':
        return appendRow_(request);
      case 'write':
        return writeRange_(request);
      case 'appendJobAlert':
        return appendJobAlert_(request);
      case 'appendVacancyAlert':
        return appendVacancyAlert_(request);
      default:
        throw new Error(`Operação não suportada: ${request.operation}`);
    }
  } catch (error) {
    return { ok: false, error: error.message };
  }
}

function appendRow_(request) {
  const config = getSheetConfig_(request.sheetKey);
  const rangeName = String(request.range || '');
  assertAllowedRange_(request.sheetKey, rangeName);
  const values = request.values;
  const parts = rangeName.split('!');
  const columnMatch = parts[1].match(/^([A-Z]+):([A-Z]+)$/);
  if (!columnMatch || !Array.isArray(values) || values.length === 0) {
    throw new Error('append requer um intervalo de colunas e uma linha não vazia.');
  }
  const firstColumn = columnToNumber_(columnMatch[1]);
  const width = columnToNumber_(columnMatch[2]) - firstColumn + 1;
  if (values.length !== width) {
    throw new Error(`append requer ${width} valores para ${rangeName}.`);
  }
  const sheet = SpreadsheetApp.openById(config.spreadsheetId).getSheetByName(parts[0]);
  if (!sheet) throw new Error(`A aba ${parts[0]} não existe.`);
  sheet.getRange(sheet.getLastRow() + 1, firstColumn, 1, width).setValues([values]);
  return { ok: true };
}

function appendJobAlert_(request) {
  const central = getSheetConfig_('alertasCentral');
  const alert = request.alert;
  if (!alert || typeof alert !== 'object') {
    throw new Error('appendJobAlert requer um objeto alert.');
  }

  const title = String(alert.title || '').trim();
  const url = String(alert.url || '').trim();
  if (!title || !/^https:\/\/www\.linkedin\.com\/jobs\/view\/\d+/.test(url)) {
    throw new Error('appendJobAlert requer título e URL válida de vaga do LinkedIn.');
  }

  const spreadsheet = SpreadsheetApp.openById(central.spreadsheetId);
  const sheet = spreadsheet.getSheetByName('Alertas LinkedIn') || spreadsheet.insertSheet('Alertas LinkedIn');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Origem', 'Cargo', 'Empresa', 'Local', 'Link', 'Detectado em']);
  }
  sheet.appendRow([
    'LinkedIn',
    title,
    String(alert.company || '').trim(),
    String(alert.location || '').trim(),
    url,
    new Date(),
  ]);
  return { ok: true };
}

function appendVacancyAlert_(request) {
  const central = getSheetConfig_('alertasCentral');
  const alert = request.alert;
  const platform = String(alert && alert.platform || '').trim();
  if (!['CIEE', 'ESPRO TAQE', 'ISBET', 'IEL Bahia', 'Start Carreiras'].includes(platform)) {
    throw new Error('appendVacancyAlert requer uma plataforma permitida.');
  }

  const title = String(alert.title || '').trim();
  const url = String(alert.url || '').trim();
  if (!title || !/^https:\/\//.test(url)) {
    throw new Error('appendVacancyAlert requer título e URL HTTPS.');
  }

  const spreadsheet = SpreadsheetApp.openById(central.spreadsheetId);
  const sheet = spreadsheet.getSheetByName('Alertas capturados') || spreadsheet.insertSheet('Alertas capturados');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Plataforma', 'Cargo', 'Empresa', 'Local', 'Link', 'Detectado em']);
  }
  sheet.appendRow([
    platform,
    title,
    String(alert.company || '').trim(),
    String(alert.location || '').trim(),
    url,
    new Date(),
  ]);
  return { ok: true };
}

function writeRange_(request) {
  const range = getAuthorizedRange_(request);
  const values = request.values;
  if (!Array.isArray(values) || !values.every(Array.isArray)) {
    throw new Error('write requer uma matriz em values.');
  }
  if (values.length !== range.getNumRows() || values[0].length !== range.getNumColumns()) {
    throw new Error('A matriz values não corresponde ao intervalo solicitado.');
  }
  range.setValues(values);
  return { ok: true };
}

function getAuthorizedRange_(request) {
  const config = getSheetConfig_(request.sheetKey);
  const rangeName = String(request.range || '');
  assertAllowedRange_(request.sheetKey, rangeName);
  const parts = rangeName.split('!');
  const sheet = SpreadsheetApp.openById(config.spreadsheetId).getSheetByName(parts[0]);
  if (!sheet) throw new Error(`A aba ${parts[0]} não existe.`);
  return sheet.getRange(parts[1]);
}

function buildConfig_(form, requireSources) {
  const previous = getConfig_();
  const sheets = {};
  Object.keys(SHEET_SPECS).forEach((key) => {
    const spreadsheetId = extractSpreadsheetId_(form[key]);
    if (!spreadsheetId) return;
    SpreadsheetApp.openById(spreadsheetId);
    sheets[key] = { spreadsheetId };
  });
  if (requireSources) {
    const missing = sourceKeys_().filter((key) => !sheets[key]);
    if (missing.length) {
      throw new Error(`Informe as seis fontes antes de criar a central: ${missing.map((key) => SHEET_SPECS[key].label).join(', ')}.`);
    }
  }
  return { version: 1, accessToken: previous.accessToken || createAccessToken_(), sheets };
}

function sourceKeys_() {
  return Object.keys(SHEET_SPECS).filter((key) => key !== 'alertasCentral');
}

function clearImportedTabs_(spreadsheet) {
  spreadsheet.getSheets().forEach((sheet) => {
    if (sheet.getName() !== 'Instruções') spreadsheet.deleteSheet(sheet);
  });
}

function recoverCentralOnlyState_() {
  const config = getConfig_();
  if (Object.keys(config.sheets).length !== 0 || !config.accessToken) return config;

  SpreadsheetApp.openById(CENTRAL_SPREADSHEET_ID);
  config.sheets.alertasCentral = { spreadsheetId: CENTRAL_SPREADSHEET_ID };
  config.lastMerge = {
    centralSpreadsheetId: CENTRAL_SPREADSHEET_ID,
    sourceKeys: sourceKeys_(),
    sourcesDeletedAt: new Date().toISOString(),
  };
  saveConfig_(config);
  return config;
}

function createDestinationName_(spreadsheet, sourceLabel, sourceTabName) {
  const base = `${sourceLabel} - ${sourceTabName}`.slice(0, 100);
  let name = base;
  let suffix = 2;
  while (spreadsheet.getSheetByName(name)) {
    name = `${base.slice(0, 96)} ${suffix++}`;
  }
  return name;
}

function getSheetConfig_(sheetKey) {
  const key = String(sheetKey || '');
  if (!Object.prototype.hasOwnProperty.call(SHEET_SPECS, key)) throw new Error(`Planilha não reconhecida: ${key}`);
  const config = getConfig_().sheets[key];
  if (!config) throw new Error(`A planilha ${SHEET_SPECS[key].label} não foi configurada.`);
  return config;
}

function assertAllowedRange_(sheetKey, rangeName) {
  if (!SHEET_SPECS[sheetKey].allowedRanges.includes(rangeName)) {
    throw new Error(`Intervalo não permitido para ${SHEET_SPECS[sheetKey].label}: ${rangeName}`);
  }
}

function requireAccessToken_(request) {
  const expected = getConfig_().accessToken;
  if (!expected || request.accessToken !== expected) throw new Error('Não autorizado. Configure o adapter e use o token de acesso atual.');
}

function getConfig_() {
  const value = PropertiesService.getScriptProperties().getProperty(CONFIG_PROPERTY);
  return value ? JSON.parse(value) : { version: 1, accessToken: '', sheets: {} };
}

function saveConfig_(config) {
  PropertiesService.getScriptProperties().setProperty(CONFIG_PROPERTY, JSON.stringify(config));
}

function createAccessToken_() {
  return [Utilities.getUuid(), Utilities.getUuid(), Utilities.getUuid(), Utilities.getUuid()].join('');
}

function extractSpreadsheetId_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : text;
}

function columnToNumber_(column) {
  return column.split('').reduce((value, character) => value * 26 + character.charCodeAt(0) - 64, 0);
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
