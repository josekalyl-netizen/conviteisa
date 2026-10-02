/**
 * Planilha do convite da Isa — recebe confirmações de presença e presentes.
 *
 * COMO LIGAR (uma vez só, uns 5 minutos):
 * 1. Crie uma planilha nova no Google Sheets (sheets.new).
 * 2. Menu Extensões > Apps Script. Apague o que estiver lá, cole ESTE código inteiro e clique em Salvar.
 * 3. Troque a SENHA abaixo por uma senha só sua (é ela que abre o painel).
 *    Em EMAIL_AVISO coloque o e-mail que deve receber um aviso a cada confirmação (ou deixe '').
 * 4. Clique em Implantar > Nova implantação > engrenagem > "App da Web".
 *    Executar como: Eu | Quem pode acessar: Qualquer pessoa > Implantar > Autorizar acesso
 *    (se aparecer "O Google não verificou este app": Avançado > Acessar... > Permitir).
 * 5. Copie a URL que termina em /exec e envie para quem cuida do convite
 *    (ela vai em CONFIG.planilhaUrl no index.html).
 *
 * Se editar este código depois: Implantar > Gerenciar implantações > lápis > Versão: Nova versão > Implantar.
 * A URL continua a mesma.
 */
const SENHA = 'troque-esta-senha';
const EMAIL_AVISO = '';   // ex.: 'familia@gmail.com' — recebe um e-mail a cada confirmação e presente

const CAMPOS = ['id','enviadoEm','nome','telefone','presenca','pessoas','acompanhantes','nomesAcompanhantes','presente','mensagem'];
const CAMPOS_PRESENTE = ['id','enviadoEm','nome','itens','valor','pagamento','mensagem','ids'];

// 1ª aba: confirmações de presença | aba "Presentes": presentes informados pelos convidados
function aba_() {
  const pl = SpreadsheetApp.getActiveSpreadsheet();
  const sh = pl.getSheetByName('Confirmações') || pl.getSheets()[0];
  if (sh.getLastRow() === 0) { sh.appendRow(CAMPOS); sh.setFrozenRows(1); sh.getRange(1, 1, 1, CAMPOS.length).setFontWeight('bold'); }
  if (sh.getName() !== 'Confirmações' && !pl.getSheetByName('Confirmações')) sh.setName('Confirmações');
  return sh;
}
function abaPresentes_() {
  const pl = SpreadsheetApp.getActiveSpreadsheet();
  const sh = pl.getSheetByName('Presentes') || pl.insertSheet('Presentes');
  if (sh.getLastRow() === 0) { sh.appendRow(CAMPOS_PRESENTE); sh.setFrozenRows(1); sh.getRange(1, 1, 1, CAMPOS_PRESENTE.length).setFontWeight('bold'); }
  // planilha criada antes da coluna "ids" (códigos dos itens da lista): acrescenta o título
  else if (String(sh.getRange(1, CAMPOS_PRESENTE.length).getValue()) !== 'ids') sh.getRange(1, CAMPOS_PRESENTE.length).setValue('ids').setFontWeight('bold');
  return sh;
}
function ler_(sh) {
  const v = sh.getDataRange().getValues();
  const cab = v.shift();
  return v.filter(l => l.some(c => c !== '')).map(l => Object.fromEntries(cab.map((c, i) => [c, l[i] instanceof Date ? l[i].toISOString() : l[i]])));
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
// Evita linha duplicada se o convite reenviar a mesma resposta (mesmo id).
function jaExiste_(sh, id) {
  if (!id || sh.getLastRow() < 2) return false;
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().some(l => String(l[0]) === String(id));
}
// Apaga a linha com esse id (usado pelo botão "apagar" do painel, protegido pela senha).
function apagar_(sh, id) {
  if (!id || sh.getLastRow() < 2) return { ok: false, erro: 'id' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]) === String(id)) { sh.deleteRow(i + 2); return { ok: true }; }
    }
    return { ok: false, erro: 'nao-encontrado' };
  } finally { lock.releaseLock(); }
}
function avisar_(assunto, linhas) {
  if (!EMAIL_AVISO) return;
  try { MailApp.sendEmail(EMAIL_AVISO, assunto, linhas.filter(Boolean).join('\n') + '\n\nPlanilha: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()); } catch (e) {}
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
      d.ids = String(d.ids || '').replace(/[^\w,-]/g, '');
      if (jaExiste_(abaPresentes_(), d.id)) return json_({ ok: true, repetido: true });
      abaPresentes_().appendRow(CAMPOS_PRESENTE.map(c => c === 'valor' ? d.valor : celula_(d[c])));
      avisar_('🎁 Presente: ' + d.nome + ' (R$ ' + d.valor.toFixed(2).replace('.', ',') + ')',
        ['Quem: ' + d.nome, 'Itens: ' + (d.itens || ''), 'Pagamento: ' + (d.pagamento === 'cartao' ? 'cartão' : 'Pix'), d.mensagem ? 'Recado: ' + d.mensagem : '',
         'Confira no extrato do banco / app do cartão.']);
      return json_({ ok: true });
    }
    d.presenca = d.presenca === 'sim' ? 'sim' : 'nao';
    d.acompanhantes = Math.max(0, Math.min(4, Number(d.acompanhantes) || 0));
    d.pessoas = d.presenca === 'sim' ? 1 + d.acompanhantes : 0;
    if (jaExiste_(aba_(), d.id)) return json_({ ok: true, repetido: true });
    aba_().appendRow(CAMPOS.map(c => celula_(d[c])));
    avisar_((d.presenca === 'sim' ? '✅ Vai: ' : '❌ Não vai: ') + d.nome + (d.presenca === 'sim' ? ' (' + d.pessoas + ' pessoa' + (d.pessoas > 1 ? 's' : '') + ')' : ''),
      ['Nome: ' + d.nome, 'WhatsApp: ' + (d.telefone || ''), d.nomesAcompanhantes ? 'Acompanhantes: ' + d.nomesAcompanhantes : '', d.mensagem ? 'Recado: ' + d.mensagem : '']);
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.acao === 'ping') return json_({ ok: true, ping: true });
  // Pública: quais itens da lista já foram dados (sem nomes nem valores), para o convite marcar "já presenteado".
  if (p.acao === 'comprados') return json_({ ok: true, comprados: ler_(abaPresentes_()).map(r => ({ ids: String(r.ids || ''), itens: String(r.itens || '') })) });
  if ((p.senha || '') !== SENHA) return json_({ ok: false, erro: 'senha' });
  if (p.acao === 'apagar') return json_(apagar_(p.tipo === 'presente' ? abaPresentes_() : aba_(), p.id));
  return json_({ ok: true, respostas: ler_(aba_()), presentes: ler_(abaPresentes_()), planilha: SpreadsheetApp.getActiveSpreadsheet().getUrl() });
}

// Opcional: rode esta função uma vez (botão ▶ Executar) para criar as abas e testar o e-mail.
function testar() {
  aba_(); abaPresentes_();
  avisar_('Teste do convite da Isa', ['Se você recebeu este e-mail, os avisos estão funcionando.']);
}
