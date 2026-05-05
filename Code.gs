// ================================================================
//  ERGO X — Google Apps Script (Code.gs)
//
//  PASSO 1: Substitua 'SEU_ID_AQUI' pelo ID da sua Planilha Google.
//           O ID fica na URL da planilha:
//           docs.google.com/spreadsheets/d/ >> ID_AQUI << /edit
//
//  PASSO 2: No Apps Script, vá em Implantar > Nova implantação:
//           Tipo: App da Web
//           Executar como: Eu (seu e-mail)
//           Quem tem acesso: Qualquer pessoa
//           → Copie a URL gerada e cole em app.js (CONFIG.API_URL)
// ================================================================

const SPREADSHEET_ID = '16Xfx3xdWNJIeZlRMwnl7VdZmGHHzE68NtG4CGq8LcXU';

const SHEET_NAMES = {
  AET: 'BD_AET',
  PA:  'BD_PA'
};

// ── Cabeçalhos esperados de cada planilha ──────────────────────
const HEADERS = {
  BD_AET: [
    'ID', 'SETOR', 'POSTO_TRABALHO', 'CRITICIDADE_ATUAL',
    'CRITICIDADE_2024', 'CRITICIDADE_2023', 'CRITICIDADE_2022',
    'CRITICIDADE_2021', 'CRITICIDADE_2020', 'CRITICIDADE_2019',
    'POSTO_GENERO', 'ATUALIZACAO', 'GERENTE', 'OBSERVACOES', 'CONDICAO_UNISSEX'
  ],
  BD_PA: [
    'ID', 'SETOR', 'POSTO_TRABALHO', 'CRITICIDADE', 'ACAO_CONTROLE',
    'CLASSIFICACAO', 'ESTIMATIVA_VALOR', 'GERENTE', 'RESPONSAVEL',
    'DATA_PREVISTA', 'DATA_CONCLUSAO', 'STATUS', 'OBSERVACOES', 'EFICACIA'
  ]
};

// ================================================================
//  ENTRY POINT
// ================================================================
function doGet(e) {
  try {
    const action   = e.parameter.action;
    const sheetKey = e.parameter.sheet;
    const name     = SHEET_NAMES[sheetKey];

    if (!name) throw new Error('Sheet inválida: ' + sheetKey);

    let result;
    if (action === 'read') {
      result = readSheet(name);

    } else if (action === 'create') {
      const data = decodePayload(e.parameter.data);
      result = createRow(name, data);

    } else if (action === 'update') {
      const rowNum = parseInt(e.parameter.rowNum, 10);
      const data   = decodePayload(e.parameter.data);
      result = updateRow(name, rowNum, data);

    } else if (action === 'delete') {
      const rowNum = parseInt(e.parameter.rowNum, 10);
      result = deleteRow(name, rowNum);

    } else {
      throw new Error('Ação desconhecida: ' + action);
    }

    return jsonOk(result);

  } catch (err) {
    return jsonErr(err.message);
  }
}

// ================================================================
//  CRUD
// ================================================================
function readSheet(name) {
  const sheet = getSheet(name);
  const vals  = sheet.getDataRange().getValues();
  if (vals.length < 2) return [];

  const headers = vals[0];
  return vals.slice(1)
    .map((row, i) => {
      const obj = { _row: i + 2 };
      headers.forEach((h, j) => {
        const v = row[j];
        obj[h] = v instanceof Date ? Utilities.formatDate(v, 'America/Sao_Paulo', 'yyyy-MM-dd') : v;
      });
      return obj;
    })
    .filter(r => r.ID || r.SETOR); // ignora linhas vazias
}

function createRow(name, data) {
  const sheet   = getSheet(name);
  const headers = getHeaders(name);

  // Gera ID baseado no número da próxima linha
  data.ID = sheet.getLastRow(); // linha atual será lastRow + 1

  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);

  return { id: data.ID, rowNum: sheet.getLastRow() };
}

function updateRow(name, rowNum, data) {
  const sheet   = getSheet(name);
  const headers = getHeaders(name);

  // Lê os valores existentes para não sobrescrever campos não enviados
  const existing = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];

  const newRow = headers.map((h, i) => {
    if (data[h] !== undefined && data[h] !== null) return data[h];
    return existing[i];
  });

  sheet.getRange(rowNum, 1, 1, newRow.length).setValues([newRow]);
  return { updated: true, rowNum };
}

function deleteRow(name, rowNum) {
  getSheet(name).deleteRow(rowNum);
  return { deleted: true, rowNum };
}

// ================================================================
//  SETUP — Cria os cabeçalhos se a planilha estiver vazia
//  Execute esta função UMA VEZ após criar a planilha Google.
// ================================================================
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  Object.entries(HEADERS).forEach(([sheetName, headers]) => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
      Logger.log('Cabeçalhos criados em: ' + sheetName);
    } else {
      Logger.log(sheetName + ' já possui dados, pulando.');
    }
  });

  Logger.log('Setup concluído!');
}

// ================================================================
//  HELPERS
// ================================================================
function getSheet(name) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Aba não encontrada: ' + name);
  return sheet;
}

function getHeaders(name) {
  const sheet = getSheet(name);
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function decodePayload(b64) {
  const bytes = Utilities.base64Decode(b64);
  const str   = Utilities.newBlob(bytes).getDataAsString('UTF-8');
  return JSON.parse(str);
}

function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
