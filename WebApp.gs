// =====================================================================
// WebApp.gs — Grupo LOS | Monitoramento Financeiro
// Compatível com GitHub Pages (CORS via ContentService)
// =====================================================================

var SPREADSHEET_ID   = '1-0xTOqOBvYyN90hRRlqaBBbgD4YA6PIBbHoZc_SRv84';
var ABA_CUSTO_SOLICIT = 'CUSTO SOLICITADO';
var ABA_DIARIAS      = 'DIARIAS CD';
var ABA_MOTORISTAS   = 'MOTORISTAS';

function responder(obj, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(obj) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET: serve HTML (quando acessado direto) ou dados ──
function doGet(e) {
  var acao = (e && e.parameter && e.parameter.acao) ? e.parameter.acao : null;
  var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;
  if (!acao) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Monitoramento Financeiro')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (acao === 'motoristas') return responder(getMotoristasList(), callback);
  if (acao === 'registros')  return responder(getRegistrosList(), callback);
  if (acao === 'diarias')    return responder(getDiariasList(), callback);
  return responder({ erro: 'Ação GET não encontrada: ' + acao }, callback);
}

// ── POST: operações de escrita ──
function doPost(e) {
  try {
    var body  = JSON.parse(e.postData.contents);
    var acao  = body.acao;
    var dados = body.dados || {};

    if (acao === 'salvarRegistro')        return responder(salvarRegistroWeb(dados));
    if (acao === 'salvarDiaria')          return responder(salvarDiariaWeb(dados));
    if (acao === 'atualizarStatus')       return responder(atualizarStatusWeb(dados));
    if (acao === 'atualizarStatusDiaria') return responder(atualizarStatusDiariaWeb(dados));
    if (acao === 'editarRegistro')        return responder(editarRegistroWeb(dados));
    if (acao === 'editarDiaria')          return responder(editarDiariaWeb(dados));
    if (acao === 'excluirRegistro')       return responder(excluirRegistro(dados));
    if (acao === 'excluirDiaria')         return responder(excluirDiaria(dados));
    if (acao === 'excluirPeriodo')        return responder(excluirPorPeriodo(dados));
    if (acao === 'cadastrarMotorista')    return responder(cadastrarMotorista(dados));
    if (acao === 'excluirMotorista')      return responder(excluirMotorista(dados));

    return responder({ erro: 'Ação POST não encontrada: ' + acao });
  } catch(err) {
    return responder({ erro: err.message });
  }
}

// ── MOTORISTAS ──
function getMotoristasList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_MOTORISTAS);
  var vals = ws.getDataRange().getValues();
  var lista = [];
  for (var i = 1; i < vals.length; i++) {
    var nome = vals[i][1];
    if (nome && String(nome).trim() !== '') lista.push({ linha: i+1, nome: String(nome).trim() });
  }
  return lista;
}

function cadastrarMotorista(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_MOTORISTAS);
  var vals = ws.getDataRange().getValues();
  var nome = String(dados.nome || '').trim().toUpperCase();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][1]).trim().toUpperCase() === nome) return { erro: 'Motorista já cadastrado!' };
  }
  var ultima = 1;
  for (var i = 1; i < vals.length; i++) { if (vals[i][1] !== '') ultima = i+1; }
  ws.getRange(ultima+1, 2).setValue(nome);
  return { sucesso: true };
}

function excluirMotorista(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_MOTORISTAS);
  ws.deleteRow(Number(dados.linha));
  return { sucesso: true };
}

// ── REGISTROS (CUSTO SOLICITADO) ──
function getRegistrosList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_CUSTO_SOLICIT);
  var vals = ws.getDataRange().getValues();
  var lista = [];
  for (var i = 1; i < vals.length; i++) {
    if (!vals[i][0]) continue;
    var dt = vals[i][4];
    var dtStr = '';
    if (dt instanceof Date) {
      var d = dt.getDate(), mo = dt.getMonth()+1, yr = dt.getFullYear();
      var h = dt.getHours(), mi = dt.getMinutes();
      dtStr = (d<10?'0':'')+d+'/'+(mo<10?'0':'')+mo+'/'+yr+' '+(h<10?'0':'')+h+':'+(mi<10?'0':'')+mi;
    } else { dtStr = String(dt); }
    lista.push({
      linha: i+1,
      motorista: String(vals[i][0]),
      remessa:   String(vals[i][1]),
      evento:    String(vals[i][2]),
      valor:     Number(vals[i][3]) || 0,
      data:      dtStr,
      status:    String(vals[i][5] || 'SELECIONAR'),
      obs:       String(vals[i][6] || '')
    });
  }
  return lista;
}

function salvarRegistroWeb(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_CUSTO_SOLICIT);
  var colA = ws.getRange('A:A').getValues();
  var ultima = 1;
  for (var i = 1; i < colA.length; i++) {
    if (String(colA[i][0]).trim() !== '') ultima = i + 1;
  }
  ultima++;
  var valor = Number(dados.valor) || 0;
  var status = valor < 0 ? 'NEGATIVAR' : 'SELECIONAR';
  ws.getRange(ultima, 1).setValue(String(dados.mot || '').toUpperCase());
  ws.getRange(ultima, 2).setValue(dados.rem || '');
  ws.getRange(ultima, 3).setValue(String(dados.evt || '').toUpperCase());
  ws.getRange(ultima, 4).setValue(valor);
  ws.getRange(ultima, 5).setValue(new Date());
  ws.getRange(ultima, 6).setValue(status);
  return { sucesso: true };
}

function atualizarStatusWeb(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_CUSTO_SOLICIT);
  var linha = Number(dados.linha);
  ws.getRange(linha, 6).setValue(dados.status || 'SELECIONAR');
  if (dados.obs !== undefined) ws.getRange(linha, 7).setValue(dados.obs || '');
  return { sucesso: true };
}

function editarRegistroWeb(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_CUSTO_SOLICIT);
  var linha = Number(dados.linha);
  var valor = Number(dados.valor) || 0;
  ws.getRange(linha, 1).setValue(String(dados.mot || '').toUpperCase());
  ws.getRange(linha, 2).setValue(dados.rem || '');
  ws.getRange(linha, 3).setValue(String(dados.evt || '').toUpperCase());
  ws.getRange(linha, 4).setValue(valor);
  if (dados.obs !== undefined) ws.getRange(linha, 7).setValue(dados.obs || '');
  return { sucesso: true };
}

function excluirRegistro(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_CUSTO_SOLICIT);
  ws.deleteRow(Number(dados.linha));
  return { sucesso: true };
}

function excluirPorPeriodo(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_CUSTO_SOLICIT);
  var vals = ws.getDataRange().getValues();
  var ini = new Date(dados.ini);
  var fim = new Date(dados.fim);
  fim.setHours(23,59,59);
  var linhasExcluir = [];
  for (var i = vals.length-1; i >= 1; i--) {
    if (!vals[i][0]) continue;
    var dt = vals[i][4];
    if (dt instanceof Date && dt >= ini && dt <= fim) linhasExcluir.push(i+1);
  }
  for (var j = 0; j < linhasExcluir.length; j++) ws.deleteRow(linhasExcluir[j]);
  return { sucesso: true, excluidos: linhasExcluir.length };
}

// ── DIÁRIAS CD ──
function getDiariasList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_DIARIAS);
  if (!ws) return [];
  var vals = ws.getDataRange().getValues();
  var lista = [];
  for (var i = 1; i < vals.length; i++) {
    if (!vals[i][0]) continue;
    var dt = vals[i][5];
    var dtStr = '';
    if (dt instanceof Date) {
      var d = dt.getDate(), mo = dt.getMonth()+1, yr = dt.getFullYear();
      var h = dt.getHours(), mi = dt.getMinutes();
      dtStr = (d<10?'0':'')+d+'/'+(mo<10?'0':'')+mo+'/'+yr+' '+(h<10?'0':'')+h+':'+(mi<10?'0':'')+mi;
    } else { dtStr = String(dt); }
    lista.push({
      linha:     i+1,
      motorista: String(vals[i][0]),
      remessa:   String(vals[i][1]),
      evento:    String(vals[i][2]),
      vd:        Number(vals[i][3]) || 0,
      vo:        Number(vals[i][4]) || 0,
      data:      dtStr,
      status:    String(vals[i][6] || 'SELECIONAR')
    });
  }
  return lista;
}

function salvarDiariaWeb(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_DIARIAS);
  var colA = ws.getRange('A:A').getValues();
  var ultima = 1;
  for (var i = 1; i < colA.length; i++) {
    if (String(colA[i][0]).trim() !== '') ultima = i + 1;
  }
  ultima++;
  var vd = Number(dados.vd) || 0;
  var vo = Number(dados.vo) || 0;
  ws.getRange(ultima, 1).setValue(String(dados.mot || '').toUpperCase());
  ws.getRange(ultima, 2).setValue(dados.rem || '');
  ws.getRange(ultima, 3).setValue(String(dados.evt || 'DIARIA CD').toUpperCase());
  ws.getRange(ultima, 4).setValue(vd);
  ws.getRange(ultima, 5).setValue(vo);
  ws.getRange(ultima, 6).setValue(new Date());
  ws.getRange(ultima, 7).setValue('SELECIONAR');
  return { sucesso: true };
}

function atualizarStatusDiariaWeb(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_DIARIAS);
  ws.getRange(Number(dados.linha), 7).setValue(dados.status || 'SELECIONAR');
  return { sucesso: true };
}

function editarDiariaWeb(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_DIARIAS);
  var linha = Number(dados.linha);
  ws.getRange(linha, 1).setValue(String(dados.mot || '').toUpperCase());
  ws.getRange(linha, 2).setValue(dados.rem || '');
  ws.getRange(linha, 4).setValue(Number(dados.vd) || 0);
  ws.getRange(linha, 5).setValue(Number(dados.vo) || 0);
  return { sucesso: true };
}

function excluirDiaria(dados) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(ABA_DIARIAS);
  ws.deleteRow(Number(dados.linha));
  return { sucesso: true };
}
