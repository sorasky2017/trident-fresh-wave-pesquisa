// Trident Fresh Wave - cadastro, pontos e trocas (planilha como banco de dados)
// Cole em Extensões > Apps Script da planilha e implante como App da Web (acesso: Qualquer pessoa).
var SENHA = 'ONDA7';      // senha do estande (só existe aqui)
var BONUS = 60;           // pontos da senha
var TZ = 'America/Sao_Paulo';
var DESCARTAVEIS = ['mailinator.com','guerrillamail.com','10minutemail.com','tempmail.com','temp-mail.org','yopmail.com','trashmail.com','throwawaymail.com','sharklasers.com','getnada.com','maildrop.cc','dispostable.com','fakeinbox.com','mohmal.com','emailondeck.com'];
var ITENS = {
  kit:      { nome: 'Kit 5 sabores',        custo: 200 },
  chiclete: { nome: 'Chiclete Trident',     custo: 120 },
  chaveiro: { nome: 'Chaveiro Fresh Wave',  custo: 60 },
  mochila:  { nome: 'Mochila Trident',      custo: 1000 },
  celular:  { nome: 'iPhone 17',            custo: 50000 },
  console:  { nome: 'PS5', custo: 75000 },
  notebook: { nome: 'Notebook',             custo: 100000 },
  pc:       { nome: 'Computador gamer',     custo: 150000 }
};
var CAD = 'Cadastros';    // Data | Nome | E-mail | Sabor | LGPD
var MOV = 'Movimentos';   // Data | E-mail | Descrição | Pontos | Código | Entregue?

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'painel') return resp(painel());
  if (p.action === 'feira') {
    var lk = LockService.getScriptLock();
    lk.waitLock(15000);
    try { return resp(feira(p)); } catch (err) { return resp({ ok: false, motivo: 'erro' }); } finally { lk.releaseLock(); }
  }
  var email = String(p.email || '').trim().toLowerCase();
  if (!/^[a-z0-9._%+\-]+@[a-z0-9\-]+(\.[a-z0-9\-]+)*\.[a-z]{2,}$/i.test(email) || /\.\./.test(email) || email.length > 120) return resp({ ok: false, motivo: 'email' });
  if (DESCARTAVEIS.indexOf(email.split('@')[1]) > -1) return resp({ ok: false, motivo: 'email' });
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    switch (p.action) {
      case 'cadastro': return resp(cadastro(email, p));
      case 'entrar':   return resp(entrar(email));
      case 'senha':    return resp(senha(email, String(p.senha || '').trim().toUpperCase()));
      case 'trocar':   return resp(trocar(email, String(p.item || '')));
      default:         return resp({ ok: false, motivo: 'acao' });
    }
  } catch (err) {
    return resp({ ok: false, motivo: 'erro' });
  } finally {
    lock.releaseLock();
  }
}

function aba(nome, cab) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(nome) || ss.insertSheet(nome);
  if (sh.getLastRow() === 0) { sh.appendRow(cab); sh.setFrozenRows(1); estilo(sh); }
  return sh;
}
function cad() { return aba(CAD, ['Data', 'Nome', 'E-mail', 'Sabor', 'LGPD']); }
function mov() { return aba(MOV, ['Data', 'E-mail', 'Descrição', 'Pontos', 'Código', 'Entregue?']); }
function limpa(s, n) { s = String(s || '').replace(/[\r\n\t]+/g, ' ').trim().slice(0, n); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function fmt(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'dd/MM/yyyy HH:mm') : String(v); }
function agora() { return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'); }

function achar(email) {
  var sh = cad(), n = sh.getLastRow();
  if (n < 2) return null;
  var v = sh.getRange(2, 1, n - 1, 5).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][2]).toLowerCase() === email) return { linha: i + 2, nome: v[i][1], sabor: v[i][3] };
  }
  return null;
}
function extrato(email) {
  var sh = mov(), n = sh.getLastRow(), saldo = 0, hist = [], usada = false;
  if (n >= 2) {
    var v = sh.getRange(2, 1, n - 1, 6).getValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][1]).toLowerCase() !== email) continue;
      var pts = Number(v[i][3]) || 0;
      saldo += pts;
      if (String(v[i][2]) === 'Senha do estande') usada = true;
      hist.unshift({ t: String(v[i][2]), d: fmt(v[i][0]), p: pts, c: String(v[i][4] || '') });
    }
  }
  return { saldo: saldo, hist: hist.slice(0, 15), usada: usada };
}
function ficha(email, c) {
  var x = extrato(email);
  return { ok: true, nome: c.nome, saldo: x.saldo, hist: x.hist, usada: x.usada };
}

function cadastro(email, p) {
  var nome = limpa(p.nome, 60), sabor = limpa(p.sabor, 40);
  if (!nome || !sabor || p.lgpd !== '1') return { ok: false, motivo: 'dados' };
  var c = achar(email);
  if (c) return { ok: true, jaExistia: true, nome: c.nome, sabor: c.sabor };
  cad().appendRow([agora(), nome, email, sabor, 'aceitou']);
  return { ok: true, jaExistia: false, nome: nome, sabor: sabor };
}
function entrar(email) {
  var c = achar(email);
  if (!c) return { ok: false, motivo: 'naocadastrado' };
  return ficha(email, c);
}
function senha(email, s) {
  var c = achar(email);
  if (!c) return { ok: false, motivo: 'naocadastrado' };
  if (extrato(email).usada) return { ok: false, motivo: 'usada' };
  if (s !== SENHA) return { ok: false, motivo: 'senha' };
  mov().appendRow([agora(), email, 'Senha do estande', BONUS, '', '']);
  return ficha(email, c);
}
function trocar(email, item) {
  var c = achar(email), it = ITENS[item];
  if (!c) return { ok: false, motivo: 'naocadastrado' };
  if (!it) return { ok: false, motivo: 'item' };
  if (extrato(email).saldo < it.custo) return { ok: false, motivo: 'saldo' };
  var cod = 'ONDA-' + Math.random().toString(36).slice(2, 6).toUpperCase();
  mov().appendRow([agora(), email, 'Troca · ' + it.nome, -it.custo, cod, 'Não']);
  var r = ficha(email, c);
  r.codigo = cod; r.item = it.nome;
  return r;
}
// Números agregados para o painel do estande (sem dados pessoais)
function painel() {
  var cache = CacheService.getScriptCache(), c = cache.get('painel');
  if (c) return JSON.parse(c);
  var cs = cad(), ms = mov(), nc = cs.getLastRow(), nm = ms.getLastRow();
  var r = { ok: true, cadastros: 0, completaram: 0, trocas: 0, entregues: 0, sabores: {}, porHora: {}, atualizado: agora() };
  if (nc >= 2) {
    var a = cs.getRange(2, 1, nc - 1, 5).getValues();
    r.cadastros = a.length;
    for (var i = 0; i < a.length; i++) {
      var sb = String(a[i][3]); r.sabores[sb] = (r.sabores[sb] || 0) + 1;
      var h = fmt(a[i][0]).slice(11, 13); if (h) r.porHora[h] = (r.porHora[h] || 0) + 1;
    }
  }
  if (nm >= 2) {
    var b = ms.getRange(2, 1, nm - 1, 6).getValues();
    for (var j = 0; j < b.length; j++) {
      var d = String(b[j][2]);
      if (d === 'Senha do estande') r.completaram++;
      if (d.indexOf('Troca') === 0) { r.trocas++; if (String(b[j][5]).toLowerCase() === 'sim') r.entregues++; }
    }
  }
  cache.put('painel', JSON.stringify(r), 8);
  return r;
}
function resp(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Sorteio de ecobags na feira (sem e-mail): aba "Sorteio" = Nº | Data | Nome completo | Curso | Avaliação do estande (1-5) | LGPD
function feira(p) {
  var nome = limpa(p.nome, 80), curso = limpa(p.curso, 60), nota = parseInt(p.avaliacao, 10);
  if (!nome || !curso || !(nota >= 1 && nota <= 5) || p.lgpd !== '1' || nome.split(' ').length < 2) return { ok: false, motivo: 'dados' };
  var sh = aba('Sorteio', ['Nº', 'Data', 'Nome completo', 'Curso', 'Avaliação do estande', 'LGPD']);
  if (String(sh.getRange(1, 5).getValue()) === 'Sabor') sh.getRange(1, 5).setValue('Avaliação do estande');
  var n = sh.getLastRow();
  var chave = (nome + '|' + curso).toLowerCase();
  if (n >= 2) {
    var v = sh.getRange(2, 1, n - 1, 6).getValues();
    for (var i = 0; i < v.length; i++) {
      if ((String(v[i][2]) + '|' + String(v[i][3])).toLowerCase() === chave) return { ok: true, jaExistia: true, nome: String(v[i][2]) };
    }
  }
  sh.appendRow([n, agora(), nome, curso, nota, 'aceitou']);
  return { ok: true, jaExistia: false, nome: nome };
}

// ---- Visual da planilha: rode "deixarBonita" uma vez (Apps Script > selecione a função > Executar) ----
var VERDE = '#00873D', VERDE_CLARO = '#EAF8EF';
function deixarBonita() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [CAD, MOV, 'Sorteio'].forEach(function (nome) {
    var sh = ss.getSheetByName(nome);
    if (sh) estilo(sh);
  });
}
function estilo(sh) {
  var nome = sh.getName(), cols = sh.getLastColumn(), rows = Math.max(sh.getMaxRows(), 2);
  if (cols < 1) return;
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 34);
  sh.getRange(1, 1, rows, cols).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sh.getRange(1, 1, 1, cols).setBackground(VERDE).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  // linhas alternadas
  sh.getBandings().forEach(function (b) { b.remove(); });
  sh.getRange(1, 1, rows, cols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(VERDE).setFirstRowColor('#FFFFFF').setSecondRowColor(VERDE_CLARO);
  sh.getRange(1, 1, rows, cols).setBorder(true, true, true, true, true, true, '#CFE9D8', SpreadsheetApp.BorderStyle.SOLID);
  sh.setTabColor('#00BF63');
  var larg = {};
  var regras = [];
  var dados = function (c) { return sh.getRange(2, c, rows - 1, 1); };
  if (nome === 'Sorteio') {
    larg = { 1: 60, 2: 130, 3: 260, 4: 220, 5: 150, 6: 90 };
    sh.getRange(2, 1, rows - 1, 2).setHorizontalAlignment('center');
    dados(5).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(12);
    dados(6).setHorizontalAlignment('center');
    regras.push(cor(dados(5), 1, 2, '#FFD9D9'), cor(dados(5), 3, 3, '#FFF3BF'), cor(dados(5), 4, 5, '#C9F0D6'));
  } else if (nome === CAD) {
    larg = { 1: 130, 2: 240, 3: 260, 4: 190, 5: 90 };
    sh.getRange(2, 1, rows - 1, 1).setHorizontalAlignment('center');
    dados(5).setHorizontalAlignment('center');
    var sabores = { 'Laranja': '#FDD9AE', 'Morango': '#F9C0C6', 'Limão': '#D9F2B0', 'Lichia': '#FBD3E2', 'Abacaxi': '#FAEFA6' };
    for (var s in sabores) regras.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith(s).setBackground(sabores[s]).setRanges([dados(4)]).build());
  } else if (nome === MOV) {
    larg = { 1: 130, 2: 260, 3: 250, 4: 90, 5: 110, 6: 100 };
    sh.getRange(2, 1, rows - 1, 1).setHorizontalAlignment('center');
    dados(4).setHorizontalAlignment('center').setFontWeight('bold');
    dados(5).setHorizontalAlignment('center');
    dados(6).setHorizontalAlignment('center').setFontWeight('bold');
    regras.push(
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#C0392B').setRanges([dados(4)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setFontColor(VERDE).setRanges([dados(4)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Não').setBackground('#FFF3BF').setRanges([dados(6)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Sim').setBackground('#C9F0D6').setRanges([dados(6)]).build());
  }
  for (var c in larg) sh.setColumnWidth(Number(c), larg[c]);
  sh.setConditionalFormatRules(regras);
  if (!sh.getFilter()) sh.getRange(1, 1, rows, cols).createFilter();
}
function cor(range, de, ate, fundo) {
  return SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(de, ate).setBackground(fundo).setRanges([range]).build();
}
