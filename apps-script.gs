/**
 * Planilha de confirmações do convite.
 * 1. Crie uma planilha no Google Sheets.
 * 2. Extensões > Apps Script > cole este código > Salvar.
 * 3. Troque a SENHA abaixo por uma senha só sua (é ela que protege a lista).
 *    Ela NÃO vai no convite.html — quem confere a senha é este script.
 * 4. Implantar > Nova implantação > tipo "App da Web"
 *    Executar como: Eu | Quem pode acessar: Qualquer pessoa > Implantar.
 * 5. Copie a URL que termina em /exec e cole em CONFIG.planilhaUrl no convite.html.
 *    (Se editar este código depois: Implantar > Gerenciar implantações > editar > Nova versão.)
 */
const SENHA = 'troque-esta-senha';
const CAMPOS = ['id','enviadoEm','nome','telefone','presenca','pessoas','acompanhantes','nomesAcompanhantes','presente','mensagem'];
const CAMPOS_PRESENTE = ['id','enviadoEm','nome','itens','valor','pagamento','mensagem'];

// 1ª aba: confirmações de presença | aba "Presentes": presentes informados pelos convidados
function aba_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(CAMPOS);
  return sh;
}
function abaPresentes_() {
  const pl = SpreadsheetApp.getActiveSpreadsheet();
  const sh = pl.getSheetByName('Presentes') || pl.insertSheet('Presentes');
  if (sh.getLastRow() === 0) sh.appendRow(CAMPOS_PRESENTE);
  return sh;
}
function ler_(sh) {
  const v = sh.getDataRange().getValues();
  const cab = v.shift();
  return v.map(l => Object.fromEntries(cab.map((c, i) => [c, l[i] instanceof Date ? l[i].toISOString() : l[i]])));
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
// Texto que começa com = + - @ vira fórmula no Sheets (ex.: telefone "+55 11...").
// O apóstrofo na frente guarda como texto puro.
function celula_(v) {
  const s = String(v == null ? '' : v).slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d || String(d.nome || '').trim().length < 3) return json_({ ok: false, erro: 'nome' });
    if (d.tipo === 'presente') {
      d.valor = Math.max(0, Math.round((Number(d.valor) || 0) * 100) / 100);
      d.pagamento = d.pagamento === 'cartao' ? 'cartao' : 'pix';
      abaPresentes_().appendRow(CAMPOS_PRESENTE.map(c => c === 'valor' ? d.valor : celula_(d[c])));
      return json_({ ok: true });
    }
    d.presenca = d.presenca === 'sim' ? 'sim' : 'nao';
    d.acompanhantes = Math.max(0, Math.min(4, Number(d.acompanhantes) || 0));
    d.pessoas = d.presenca === 'sim' ? 1 + d.acompanhantes : 0;
    aba_().appendRow(CAMPOS.map(c => celula_(d[c])));
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}

function doGet(e) {
  if ((e.parameter.senha || '') !== SENHA) return json_({ ok: false, erro: 'senha' });
  return json_({ ok: true, respostas: ler_(aba_()), presentes: ler_(abaPresentes_()) });
}
