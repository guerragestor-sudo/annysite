/**
 * Cadastros do site da Anny (Joy Lapa) → planilha Google + aviso para a Anny.
 * Cole este código em: Planilha > Extensões > Apps Script. Passo a passo no LEIA-ME.txt.
 */

/* ============ EDITE AQUI ============ */
const CONFIG = {
  emailAnny: '',          // e-mail que recebe o aviso de cada novo lead (vazio = sem e-mail)
  callmebotApiKey: '',    // chave do CallMeBot para aviso no WhatsApp (vazio = sem WhatsApp)
  whatsappAnny: '5511989686074',
  aba: 'Leads',
};
/* ==================================== */

const COLUNAS = ['Data', 'Nome', 'WhatsApp', 'Renda da família', 'Planta', 'Chamar no WhatsApp',
  'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'Página'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const nome = limpa(d.nome, 80);
    const tel = String(d.whatsapp || '').replace(/\D/g, '').slice(0, 11);
    if (nome.length < 2 || !/^\d{10,11}$/.test(tel)) return resposta('invalido');

    const telFmt = tel.length === 11 ? `(${tel.slice(0, 2)}) ${tel.slice(2, 7)}-${tel.slice(7)}`
                                     : `(${tel.slice(0, 2)}) ${tel.slice(2, 6)}-${tel.slice(6)}`;
    const msgCliente = `Olá, ${nome.split(' ')[0]}! Aqui é a Anny, do Joy Lapa. Recebi seu cadastro e já vou preparar sua simulação.`;
    const linkCliente = `https://wa.me/55${tel}?text=${encodeURIComponent(msgCliente)}`;
    const renda = limpa(d.renda, 40), planta = limpa(d.planta, 60);

    aba().appendRow([new Date(), nome, telFmt, renda, planta, linkCliente,
      limpa(d.utm_source, 80), limpa(d.utm_medium, 80), limpa(d.utm_campaign, 120),
      limpa(d.utm_content, 120), limpa(d.pagina, 300)]);

    const resumo = `Novo lead do Joy Lapa\n\nNome: ${nome}\nWhatsApp: ${telFmt}\nRenda: ${renda}\nPlanta: ${planta}` +
      (d.utm_campaign ? `\nCampanha: ${limpa(d.utm_campaign, 120)}` : '');

    if (CONFIG.emailAnny) {
      MailApp.sendEmail({
        to: CONFIG.emailAnny,
        subject: `Novo lead Joy Lapa: ${nome}`,
        body: `${resumo}\n\nChamar agora: ${linkCliente}\nPlanilha: ${SpreadsheetApp.getActive().getUrl()}`,
      });
    }
    if (CONFIG.callmebotApiKey) {
      UrlFetchApp.fetch('https://api.callmebot.com/whatsapp.php?phone=' + CONFIG.whatsappAnny +
        '&apikey=' + encodeURIComponent(CONFIG.callmebotApiKey) +
        '&text=' + encodeURIComponent(`${resumo}\n\nChamar: https://wa.me/55${tel}`), { muteHttpExceptions: true });
    }
    return resposta('ok');
  } catch (err) {
    console.error(err);
    return resposta('erro');
  } finally {
    lock.releaseLock();
  }
}

function aba() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.aba);
  if (!sh) sh = ss.insertSheet(CONFIG.aba);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUNAS);
    sh.getRange(1, 1, 1, COLUNAS.length).setFontWeight('bold').setBackground('#1F1633').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  }
  return sh;
}

// Impede fórmulas injetadas pelo formulário (=, +, -, @ no começo)
function limpa(v, max) {
  const s = String(v || '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function resposta(status) {
  return ContentService.createTextOutput(JSON.stringify({ status })).setMimeType(ContentService.MimeType.JSON);
}

// Rode esta função uma vez pelo editor para autorizar e-mail/planilha e testar o aviso.
function testar() {
  doPost({ postData: { contents: JSON.stringify({
    nome: 'Teste do site', whatsapp: '11999999999', renda: 'R$ 3.200 a R$ 5.000', planta: 'Select, 2 dorms',
  }) } });
}
