/**
 * Álbum da festa da Isa — recebe as fotos e vídeos que os convidados enviam pela página /fotos.
 * Os arquivos vão para a pasta "Fotos da festa – convidados" no Google Drive da conta onde este script foi criado.
 *
 * COMO LIGAR (uma vez, uns 5 minutos), logado na conta Google da Isa:
 * 1. Acesse script.google.com > Novo projeto. Apague o que estiver lá, cole ESTE código e clique em Salvar.
 * 2. Na barra de cima escolha a função "autorizar" e clique em Executar > Revisar permissões > escolha a conta
 *    (se aparecer "O Google não verificou este app": Avançado > Acessar... > Permitir).
 *    Isso cria a pasta no Drive.
 * 3. Implantar > Nova implantação > engrenagem > "App da Web"
 *    Executar como: Eu | Quem pode acessar: Qualquer pessoa > Implantar.
 * 4. Copie a URL que termina em /exec e envie para quem cuida do convite.
 */
const PASTA = 'Fotos da festa – convidados';
const ORIGEM = 'https://isaespindolaconvite.github.io';   // site que pode enviar arquivos
const LIMITE_MB = 2048;                                     // tamanho máximo por arquivo

function pasta_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PASTA');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(PASTA);
  props.setProperty('PASTA', f.getId());
  return f;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Abre uma sessão de envio no Drive para UM arquivo. O arquivo sobe direto do celular do convidado
// para o Drive; o script só cria a sessão (não vê nem guarda a senha/token).
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.acao !== 'iniciar') return json_({ ok: true, ping: true });
  const tipo = String(p.tipo || '');
  if (!/^(image|video)\//.test(tipo)) return json_({ ok: false, erro: 'tipo' });
  const tam = Number(p.tamanho) || 0;
  if (tam <= 0 || tam > LIMITE_MB * 1048576) return json_({ ok: false, erro: 'tamanho' });
  const quem = String(p.quem || '').replace(/[^\p{L}\p{N} .'-]/gu, '').trim().slice(0, 40) || 'Convidado';
  const original = String(p.nome || 'arquivo').replace(/[\\/:*?"<>|]/g, '').slice(-80);
  const quando = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd-MM HH.mm.ss');
  const meta = { name: quando + ' – ' + quem + ' – ' + original, parents: [pasta_().getId()], description: 'Enviado por ' + quem };
  const r = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable', {
    method: 'post', contentType: 'application/json; charset=UTF-8', payload: JSON.stringify(meta), muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': tipo, 'X-Upload-Content-Length': String(tam), Origin: ORIGEM }
  });
  const h = r.getAllHeaders();
  const url = h.Location || h.location;
  if (r.getResponseCode() !== 200 || !url) return json_({ ok: false, erro: 'drive', detalhe: r.getContentText().slice(0, 200) });
  return json_({ ok: true, url: url });
}

// Rode uma vez (passo 2): pede as permissões e cria a pasta. O link da pasta aparece no "Registro de execução".
function autorizar() {
  const f = pasta_();
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/about?fields=user', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  Logger.log('Pasta pronta: ' + f.getUrl());
}
