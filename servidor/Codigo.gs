// Trident Fresh Wave - controle da senha do estande
// Cole este código em script.google.com (projeto ligado a uma planilha) e publique como app da web.
var SENHA = 'ONDA7';       // senha do estande (fica só aqui, não aparece na página)
var ABA = 'Resgates';

function doGet(e) {
  var p = (e && e.parameter) || {};
  var email = String(p.email || '').trim().toLowerCase();
  var senha = String(p.senha || '').trim().toUpperCase();
  return resp(checar(email, senha));
}

function checar(email, senha) {
  if (!/^\S+@\S+\.\S+$/.test(email)) return { ok: false, motivo: 'email' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ABA) || ss.insertSheet(ABA);
    if (sh.getLastRow() === 0) sh.appendRow(['Data', 'E-mail', 'Retirou brinde?']);
    var emails = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues() : [];
    for (var i = 0; i < emails.length; i++) {
      if (String(emails[i][0]).toLowerCase() === email) return { ok: false, motivo: 'usada' };
    }
    if (senha !== SENHA) return { ok: false, motivo: 'senha' };
    sh.appendRow([new Date(), email, '']);
    return { ok: true, pontos: 60 };
  } finally {
    lock.releaseLock();
  }
}

function resp(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
