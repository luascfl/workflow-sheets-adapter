const CONFIG_PROPERTY = 'WORKFLOW_SHEETS_ADAPTER_CONFIG';

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
  alertasIsbet: {
    label: 'ISBET',
    allowedRanges: ['Alertas!A:A', 'Alertas!A:B'],
  },
  alertasIel: {
    label: 'IEL Bahia',
    allowedRanges: ['Alertas!A:A', 'Alertas!A:B'],
  },
  alertasStartCarreiras: {
    label: 'Start Carreiras',
    allowedRanges: ['Alertas!A:A', 'Alertas!A:B'],
  },
  alertasLinkedin: {
    label: 'LinkedIn',
    allowedRanges: ["Alertas!K3:K"],
  },
});

function doGet(event) {
  if (event.parameter && event.parameter.ui === 'setup') {
    return HtmlService.createHtmlOutputFromFile('Setup')
      .setTitle('Workflow Sheets Adapter');
  }
  return jsonResponse_(handleRequest_(event.parameter || {}));
}

function doPost(event) {
  const request = event.postData && event.postData.contents
    ? JSON.parse(event.postData.contents)
    : {};
  return jsonResponse_(handleRequest_(request));
}

function getSetupState() {
  const config = getConfig_();
  return {
    endpoint: ScriptApp.getService().getUrl(),
    accessToken: config.accessToken || '',
    sheets: Object.keys(SHEET_SPECS).map((key) => ({
      key,
      label: SHEET_SPECS[key].label,
      configured: Boolean(config.sheets[key]),
      spreadsheetUrl: config.sheets[key]
        ? `https://docs.google.com/spreadsheets/d/${config.sheets[key].spreadsheetId}/edit`
        : '',
    })),
  };
}

function saveSetupState(form) {
  const config = getConfig_();
  const sheets = {};

  Object.keys(SHEET_SPECS).forEach((key) => {
    const spreadsheetId = extractSpreadsheetId_(form[key]);
    if (!spreadsheetId) {
      return;
    }
    SpreadsheetApp.openById(spreadsheetId);
    sheets[key] = { spreadsheetId };
  });

  if (Object.keys(sheets).length === 0) {
    throw new Error('Informe ao menos uma planilha.');
  }

  const nextConfig = {
    version: 1,
    accessToken: config.accessToken || createAccessToken_(),
    sheets,
  };
  saveConfig_(nextConfig);
  return getSetupState();
}

function handleRequest_(request) {
  try {
    requireAccessToken_(request);
    const operation = String(request.operation || 'status');

    switch (operation) {
      case 'status':
        return {
          ok: true,
          configuredSheets: Object.keys(getConfig_().sheets),
        };
      case 'read':
        return readRange_(request);
      case 'append':
        return appendRow_(request);
      case 'write':
        return writeRange_(request);
      default:
        throw new Error(`Operação não suportada: ${operation}`);
    }
  } catch (error) {
    return { ok: false, error: error.message };
  }
}

function readRange_(request) {
  const range = getAuthorizedRange_(request);
  return {
    ok: true,
    values: range.getValues(),
  };
}

function appendRow_(request) {
  const config = getSheetConfig_(request.sheetKey);
  const rangeName = String(request.range || '');
  assertAllowedRange_(request.sheetKey, rangeName);
  const values = request.values;

  if (!Array.isArray(values) || values.length === 0) {
    throw new Error('append requer uma linha não vazia em values.');
  }

  const parts = rangeName.split('!');
  const columnMatch = parts[1].match(/^([A-Z]+):([A-Z]+)$/);
  if (!columnMatch) {
    throw new Error('append requer um intervalo de colunas, por exemplo Alertas!A:B.');
  }

  const firstColumn = columnToNumber_(columnMatch[1]);
  const width = columnToNumber_(columnMatch[2]) - firstColumn + 1;
  if (values.length !== width) {
    throw new Error(`append requer ${width} valores para ${rangeName}.`);
  }

  const sheet = SpreadsheetApp.openById(config.spreadsheetId).getSheetByName(parts[0]);
  if (!sheet) {
    throw new Error(`A aba ${parts[0]} não existe.`);
  }

  sheet.getRange(sheet.getLastRow() + 1, firstColumn, 1, width).setValues([values]);
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
  if (!sheet) {
    throw new Error(`A aba ${parts[0]} não existe.`);
  }
  return sheet.getRange(parts[1]);
}

function getSheetConfig_(sheetKey) {
  const key = String(sheetKey || '');
  if (!Object.prototype.hasOwnProperty.call(SHEET_SPECS, key)) {
    throw new Error(`Planilha não reconhecida: ${key}`);
  }
  const sheetConfig = getConfig_().sheets[key];
  if (!sheetConfig) {
    throw new Error(`A planilha ${SHEET_SPECS[key].label} não foi configurada.`);
  }
  return sheetConfig;
}

function assertAllowedRange_(sheetKey, rangeName) {
  if (!SHEET_SPECS[sheetKey].allowedRanges.includes(rangeName)) {
    throw new Error(`Intervalo não permitido para ${SHEET_SPECS[sheetKey].label}: ${rangeName}`);
  }
}

function requireAccessToken_(request) {
  const expectedToken = getConfig_().accessToken;
  if (!expectedToken || request.accessToken !== expectedToken) {
    throw new Error('Não autorizado. Configure o adapter e use o token de acesso atual.');
  }
}

function getConfig_() {
  const value = PropertiesService.getScriptProperties().getProperty(CONFIG_PROPERTY);
  return value
    ? JSON.parse(value)
    : { version: 1, accessToken: '', sheets: {} };
}

function saveConfig_(config) {
  PropertiesService.getScriptProperties().setProperty(CONFIG_PROPERTY, JSON.stringify(config));
}

function createAccessToken_() {
  return [
    Utilities.getUuid(),
    Utilities.getUuid(),
    Utilities.getUuid(),
    Utilities.getUuid(),
  ].join('');
}

function extractSpreadsheetId_(value) {
  const text = String(value || '').trim();
  if (!text) {
    return '';
  }
  const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : text;
}

function columnToNumber_(column) {
  return column.split('').reduce((value, character) => value * 26 + character.charCodeAt(0) - 64, 0);
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
