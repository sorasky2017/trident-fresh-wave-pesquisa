// Trident Fresh Wave - cadastro, pontos e trocas (planilha como banco de dados)
// Cole em Extensões > Apps Script da planilha e implante como App da Web (acesso: Qualquer pessoa).
var SENHA = 'ONDA7';      // senha do estande (só existe aqui)
var BONUS = 60;           // pontos da senha
var ADMIN_PIN = 'TROQUE123';  // PIN da página de sorteio (admin). Troque antes da feira.
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
  if (p.action === 'ultimo') return resp(ultimo());
  if (p.action === 'sortear' || p.action === 'ganhadores' || p.action === 'ausente' || p.action === 'presente') {
    var lk2 = LockService.getScriptLock();
    lk2.waitLock(15000);
    try {
      if (String(p.pin || '') !== ADMIN_PIN) return resp({ ok: false, motivo: 'pin' });
      return resp(p.action === 'sortear' ? sortear(p.qtd) : p.action === 'ausente' ? ausente(p.n) : p.action === 'presente' ? presente(p.n) : lista());
    } catch (err) { return resp({ ok: false, motivo: 'erro' }); } finally { lk2.releaseLock(); }
  }
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
  var tel = telefone(p.telefone);
  if (p.telefone && !tel) return { ok: false, motivo: 'telefone' };
  var sh = aba('Sorteio', ['Nº', 'Data', 'Nome completo', 'Curso', 'Avaliação do estande', 'LGPD', 'Código', 'Telefone']);
  if (String(sh.getRange(1, 5).getValue()) === 'Sabor') sh.getRange(1, 5).setValue('Avaliação do estande');
  if (String(sh.getRange(1, 7).getValue()) === '') sh.getRange(1, 7).setValue('Código');
  if (String(sh.getRange(1, 8).getValue()) === '') { sh.getRange(1, 8).setValue('Telefone'); estilo(sh); }
  var n = sh.getLastRow();
  var chave = (nome + '|' + curso).toLowerCase();
  var usados = {};
  if (n >= 2) {
    var v = sh.getRange(2, 1, n - 1, 7).getValues();
    for (var i = 0; i < v.length; i++) {
      if (v[i][6]) usados[String(v[i][6])] = 1;
      // mesmo nome e curso: não devolve o código (só quem se cadastrou primeiro o tem)
      if ((String(v[i][2]) + '|' + String(v[i][3])).toLowerCase() === chave) return { ok: true, jaExistia: true, nome: String(v[i][2]) };
    }
  }
  var cod;
  do { cod = novoCodigo(); } while (usados[cod]);
  var maxN = 0;
  if (n >= 2) sh.getRange(2, 1, n - 1, 1).getValues().forEach(function (r) { maxN = Math.max(maxN, Number(r[0]) || 0); });
  sh.appendRow([maxN + 1, agora(), nome, curso, nota, 'aceitou', cod, tel]);
  return { ok: true, jaExistia: false, nome: nome, codigo: cod };
}

// ---- Visual da planilha: rode "deixarBonita" uma vez (Apps Script > selecione a função > Executar) ----
var VERDE = '#00873D', VERDE_CLARO = '#EAF8EF';
function deixarBonita() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [CAD, MOV, 'Sorteio', 'Ganhadores'].forEach(function (nome) {
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
    larg = { 1: 60, 2: 130, 3: 260, 4: 220, 5: 150, 6: 90, 7: 120, 8: 150 };
    dados(7).setHorizontalAlignment('center').setFontWeight('bold');
    dados(8).setHorizontalAlignment('center');
    sh.getRange(2, 1, rows - 1, 2).setHorizontalAlignment('center');
    dados(5).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(12);
    dados(6).setHorizontalAlignment('center');
    regras.push(cor(dados(5), 1, 2, '#FFD9D9'), cor(dados(5), 3, 3, '#FFF3BF'), cor(dados(5), 4, 5, '#C9F0D6'));
  } else if (nome === 'Ganhadores') {
    larg = { 1: 130, 2: 60, 3: 260, 4: 220, 5: 150, 6: 110, 7: 120, 8: 150 };
    dados(7).setHorizontalAlignment('center').setFontWeight('bold');
    dados(8).setHorizontalAlignment('center');
    sh.hideColumns(9);
    sh.getRange(2, 1, rows - 1, 2).setHorizontalAlignment('center');
    dados(5).setHorizontalAlignment('center');
    dados(6).setHorizontalAlignment('center').setFontWeight('bold');
    regras.push(
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Ganhou').setBackground('#C9F0D6').setRanges([dados(6)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Presente').setBackground('#8FE0AE').setRanges([dados(6)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Ausente').setBackground('#FFD9D9').setRanges([dados(6)]).build());
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

// ---- Sorteio (página admin): sorteia no servidor entre quem ainda não ganhou ----
// Aba "Ganhadores" = Data | Nº | Nome completo | Curso | Avaliação | Situação (Ganhou / Presente / Ausente) | Código | Telefone | Registro (ms)
function gan() {
  var sh = aba('Ganhadores', ['Data', 'Nº', 'Nome completo', 'Curso', 'Avaliação', 'Situação', 'Código', 'Telefone', 'Registro (ms)']);
  if (String(sh.getRange(1, 7).getValue()) === '') sh.getRange(1, 7).setValue('Código');
  if (String(sh.getRange(1, 8).getValue()) === '') sh.getRange(1, 8).setValue('Telefone');
  if (String(sh.getRange(1, 9).getValue()) === '') { sh.getRange(1, 9).setValue('Registro (ms)'); estilo(sh); }
  return sh;
}
function lista() {
  var sh = gan(), n = sh.getLastRow(), out = [];
  if (n >= 2) {
    var v = sh.getRange(2, 1, n - 1, 9).getValues();
    for (var i = 0; i < v.length; i++) out.push({ d: fmt(v[i][0]), n: v[i][1], nome: String(v[i][2]), curso: String(v[i][3]), s: String(v[i][5]), c: String(v[i][6] || ''), t: String(v[i][7] || ''), ts: Number(v[i][8]) || 0 });
  }
  var total = Math.max(aba('Sorteio', ['Nº', 'Data', 'Nome completo', 'Curso', 'Avaliação do estande', 'LGPD', 'Código', 'Telefone']).getLastRow() - 1, 0);
  return { ok: true, ganhadores: out.reverse(), total: total, restantes: Math.max(total - out.length, 0) };
}
// Sorteia "qtd" pessoas (1 a 5) de uma vez, entre quem ainda não está em Ganhadores
function sortear(qtd) {
  qtd = Math.max(1, Math.min(5, parseInt(qtd, 10) || 1));
  var sh = aba('Sorteio', ['Nº', 'Data', 'Nome completo', 'Curso', 'Avaliação do estande', 'LGPD', 'Código', 'Telefone']), n = sh.getLastRow();
  if (n < 2) return { ok: false, motivo: 'vazio' };
  var g = gan(), ng = g.getLastRow(), usados = {};
  if (ng >= 2) g.getRange(2, 2, ng - 1, 1).getValues().forEach(function (r) { usados[String(r[0])] = 1; });
  var v = sh.getRange(2, 1, n - 1, 8).getValues(), pool = [];
  for (var i = 0; i < v.length; i++) if (!usados[String(v[i][0])]) pool.push(v[i]);
  if (!pool.length) return { ok: false, motivo: 'acabou' };
  var escolhidos = [], ts = Date.now();
  while (escolhidos.length < qtd && pool.length) escolhidos.push(pool.splice(Math.floor(Math.random() * pool.length), 1)[0]);
  var linhas = escolhidos.map(function (w) { return [agora(), w[0], w[2], w[3], w[4], 'Ganhou', String(w[6] || ''), String(w[7] || ''), ts]; });
  g.getRange(ng + 1, 1, linhas.length, 9).setValues(linhas);
  var r = lista();
  return { ok: true, pedidos: qtd, ganhadores: escolhidos.map(function (w) { return { n: w[0], nome: String(w[2]), curso: String(w[3]), codigo: String(w[6] || ''), telefone: String(w[7] || '') }; }), total: r.total, restantes: r.restantes };
}
// Marca um ganhador (pelo Nº) como ausente ou presente. Ausente continua fora de novos sorteios.
function marcar(num, situacao) {
  var g = gan(), n = g.getLastRow();
  if (n >= 2) {
    var v = g.getRange(2, 2, n - 1, 1).getValues();
    for (var i = v.length - 1; i >= 0; i--) {
      if (String(v[i][0]) === String(num)) {
        var c = g.getRange(i + 2, 6);
        if (String(c.getValue()) === 'Ganhou') c.setValue(situacao);
        break;
      }
    }
  }
  return lista();
}
function ausente(num) { return marcar(num, 'Ausente'); }
function presente(num) { return marcar(num, 'Presente'); }

// Painel público do sorteio: ganhadores recentes (só nome e curso) com a "idade" de cada sorteio, para o painel contar o tempo
function ultimo() {
  var r = lista(), now = Date.now(), recentes = [];
  for (var i = 0; i < r.ganhadores.length && recentes.length < 14; i++) {
    var x = r.ganhadores[i];
    if (x.s === 'Ausente') continue;
    recentes.push({ chave: x.n + '|' + x.ts, nome: x.nome, curso: x.curso, presente: x.s === 'Presente', idade: x.ts ? now - x.ts : 1e9 });
  }
  return { ok: true, total: r.total, restantes: r.restantes, recentes: recentes };
}

// Código aleatório de 5 caracteres (sem 0/O/1/I/L) no formato FW-K7M2Q
function novoCodigo() {
  var a = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789', s = '';
  for (var i = 0; i < 5; i++) s += a.charAt(Math.floor(Math.random() * a.length));
  return 'FW-' + s;
}

// Telefone brasileiro opcional: aceita 10 ou 11 dígitos (com DDD), guarda como texto "(11) 91234-5678"
function telefone(t) {
  var d = String(t || '').replace(/\D/g, '');
  if (d.length === 13 && d.indexOf('55') === 0) d = d.slice(2);
  if (d.length !== 10 && d.length !== 11) return '';
  var corpo = d.length === 11 ? d.slice(2, 7) + '-' + d.slice(7) : d.slice(2, 6) + '-' + d.slice(6);
  return "(" + d.slice(0, 2) + ") " + corpo;
}
