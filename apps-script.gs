/**
 * Planilha de confirmações do convite.
 * 1. Crie uma planilha no Google Sheets.
 * 2. Extensões > Apps Script > cole este código > Salvar.
 * 3. Troque a SENHA abaixo (é ela que protege a lista).
 * 4. Implantar > Nova implantação > tipo "App da Web"
 *    Executar como: Eu | Quem pode acessar: Qualquer pessoa > Implantar.
 * 5. Copie a URL que termina em /exec e cole em CONFIG.planilhaUrl no convite.html.
 */
const SENHA = '2811';
const CAMPOS = ['id','enviadoEm','nome','telefone','presenca','pessoas','acompanhantes','nomesAcompanhantes','presente','mensagem'];

function aba_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(CAMPOS);
  return sh;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    aba_().appendRow(CAMPOS.map(c => String(d[c] == null ? '' : d[c]).slice(0, 500)));
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}

function doGet(e) {
  if ((e.parameter.senha || '') !== SENHA) return json_({ ok: false, erro: 'senha' });
  const v = aba_().getDataRange().getValues();
  const cab = v.shift();
  const respostas = v.map(l => Object.fromEntries(cab.map((c, i) => [c, l[i] instanceof Date ? l[i].toISOString() : l[i]])));
  return json_({ ok: true, respostas });
}
