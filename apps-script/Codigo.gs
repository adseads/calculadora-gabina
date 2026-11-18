/**
 * Gabina 3D — Estoque
 * Cole este código em: Planilha Google → Extensões → Apps Script.
 * Depois: Implantar → Nova implantação → App da Web
 *   Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 *
 * Senha usada na página para alterar o estoque (troque se quiser):
 */
const SENHA = '1234';

const ABAS = {
  Filamentos: ['ID', 'Material', 'Cor', 'Acabamento', 'Cor (hex)', 'Rolos fechados', 'Rolo aberto (g)', 'Peso do rolo (g)', 'Preço (R$/kg)', 'Total (g)', 'Situação', 'Atualizado em'],
  Insumos: ['ID', 'Item', 'Categoria', 'Quantidade', 'Mínimo', 'Unidade', 'Custo unitário', 'Situação', 'Atualizado em'],
  'Histórico': ['Data', 'O que foi impresso', 'Material', 'Cor', 'Gramas', 'Custo (R$)', 'Descontou do estoque?']
};

const FILAMENTOS_INICIAIS = [
  ['PLA', 'Dourado', 'Silk', '#D4AF37', 1],
  ['PLA', 'Vermelho', 'Silk', '#C62828', 1],
  ['PLA', 'Chocolate', 'Silk', '#5D3A1A', 1],
  ['PLA', 'Roxo', 'Básico', '#6A1B9A', 2],
  ['PLA', 'Roxo bebê', 'Básico', '#C8A2C8', 1],
  ['PLA', 'Rosa', 'Básico', '#E91E63', 1],
  ['PLA', 'Rosa bebê', 'Básico', '#F8BBD0', 1],
  ['PLA', 'Azul bebê', 'Básico', '#A7C7E7', 1],
  ['PLA', 'Azul céu', 'Básico', '#4FC3F7', 1],
  ['PLA', 'Verde', 'Básico', '#2E7D32', 1],
  ['PLA', 'Verde musgo', 'Básico', '#6B7B3A', 1],
  ['PLA', 'Cor de pele', 'Básico', '#F1C9A5', 1],
  ['PLA', 'Branco', 'Básico', '#FFFFFF', 1],
  ['PLA', 'Preto', 'Básico', '#111111', 1],
  ['PLA', 'Cinza', 'Básico', '#9E9E9E', 1],
  ['PLA', 'Amarelo, azul e rosa', 'Multicolor', 'multi', 1],
  ['PETG', 'Transparente', 'Transparente', '#E0F2F1', 1],
  ['PETG', 'Roxo transparente', 'Transparente', '#B39DDB', 1]
];

/* ---------- Preparação (roda sozinha na primeira vez) ---------- */
function preparar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(ABAS).forEach(function (nome) {
    let aba = ss.getSheetByName(nome);
    if (!aba) aba = ss.insertSheet(nome);
    if (aba.getLastRow() === 0) {
      aba.getRange(1, 1, 1, ABAS[nome].length).setValues([ABAS[nome]])
        .setFontWeight('bold').setBackground('#fde9e1');
      aba.setFrozenRows(1);
    }
  });
  const fil = ss.getSheetByName('Filamentos');
  if (fil.getLastRow() === 1) {
    FILAMENTOS_INICIAIS.forEach(function (f) {
      gravarFilamento_(fil, { id: novoId_(), material: f[0], cor: f[1], acabamento: f[2], hex: f[3], fechados: f[4], aberto: 0, pesoRolo: 1000, preco: 100 }, null);
    });
  }
  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
}

/* ---------- Leitura ---------- */
function doGet() {
  preparar();
  return json_({ ok: true, filamentos: lerFilamentos_(), insumos: lerInsumos_(), historico: lerHistorico_() });
}

/* ---------- Alterações ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    preparar();
    const p = JSON.parse(e.postData.contents || '{}');
    if (String(p.senha) !== SENHA) return json_({ ok: false, erro: 'Senha incorreta.' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (p.aba === 'Filamentos') {
      const aba = ss.getSheetByName('Filamentos');
      if (p.acao === 'adicionar') {
        const f = p.item || {};
        f.id = novoId_();
        gravarFilamento_(aba, f, null);
      } else {
        const linha = acharLinha_(aba, p.id);
        if (!linha) return json_({ ok: false, erro: 'Item não encontrado.' });
        const f = filamentoDaLinha_(aba.getRange(linha, 1, 1, ABAS.Filamentos.length).getValues()[0]);
        if (p.acao === 'usar') {
          let g = Number(p.gramas) || 0;
          if (g <= 0) return json_({ ok: false, erro: 'Informe as gramas usadas.' });
          const descontar = p.descontar !== false;
          ss.getSheetByName('Histórico').appendRow([p.data ? new Date(p.data + 'T12:00:00') : new Date(), p.peca || '', f.material, f.cor, g, Math.round(g / 1000 * f.preco * 100) / 100, descontar ? 'Sim' : 'Não']);
          if (!descontar) return doGet();
          if (g > f.aberto + f.fechados * f.pesoRolo) return json_({ ok: false, erro: 'Não há filamento suficiente dessa cor no estoque.' });
          while (g > f.aberto) { g -= f.aberto; f.fechados -= 1; f.aberto = f.pesoRolo; }
          f.aberto = Math.round((f.aberto - g) * 10) / 10;
        } else if (p.acao === 'comprar') {
          f.fechados += Number(p.qtd) || 1;
          if (p.preco) f.preco = Number(p.preco);
        } else if (p.acao === 'editar') {
          Object.assign(f, p.item || {}, { id: f.id });
        } else if (p.acao === 'remover') {
          aba.deleteRow(linha);
          return doGet();
        }
        gravarFilamento_(aba, f, linha);
      }
    }

    if (p.aba === 'Insumos') {
      const aba = ss.getSheetByName('Insumos');
      if (p.acao === 'adicionar') {
        const i = p.item || {};
        i.id = novoId_();
        gravarInsumo_(aba, i, null);
      } else {
        const linha = acharLinha_(aba, p.id);
        if (!linha) return json_({ ok: false, erro: 'Item não encontrado.' });
        const i = insumoDaLinha_(aba.getRange(linha, 1, 1, ABAS.Insumos.length).getValues()[0]);
        if (p.acao === 'mover') {
          i.qtd = Math.max(0, i.qtd + (Number(p.delta) || 0));
        } else if (p.acao === 'editar') {
          Object.assign(i, p.item || {}, { id: i.id });
        } else if (p.acao === 'remover') {
          aba.deleteRow(linha);
          return doGet();
        }
        gravarInsumo_(aba, i, linha);
      }
    }
    return doGet();
  } catch (err) {
    return json_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Auxiliares ---------- */
function situacaoFil_(total) { return total < 250 ? 'Comprar' : (total < 500 ? 'Acabando' : 'OK'); }

function gravarFilamento_(aba, f, linha) {
  const fechados = Number(f.fechados) || 0, aberto = Number(f.aberto) || 0, peso = Number(f.pesoRolo) || 1000;
  const total = fechados * peso + aberto;
  const valores = [[f.id, f.material || 'PLA', f.cor || '', f.acabamento || 'Básico', f.hex || '#cccccc', fechados, aberto, peso, Number(f.preco) || 100, total, situacaoFil_(total), new Date()]];
  if (linha) aba.getRange(linha, 1, 1, valores[0].length).setValues(valores);
  else aba.appendRow(valores[0]);
}

function gravarInsumo_(aba, i, linha) {
  const qtd = Number(i.qtd) || 0, min = Number(i.minimo) || 0;
  const valores = [[i.id, i.item || '', i.categoria || 'Outro', qtd, min, i.unidade || 'un.', Number(i.custo) || 0, qtd <= min ? 'Comprar' : 'OK', new Date()]];
  if (linha) aba.getRange(linha, 1, 1, valores[0].length).setValues(valores);
  else aba.appendRow(valores[0]);
}

function filamentoDaLinha_(r) {
  return { id: String(r[0]), material: r[1], cor: r[2], acabamento: r[3], hex: r[4], fechados: Number(r[5]) || 0, aberto: Number(r[6]) || 0, pesoRolo: Number(r[7]) || 1000, preco: Number(r[8]) || 0, total: Number(r[9]) || 0, situacao: r[10], atualizado: r[11] };
}
function insumoDaLinha_(r) {
  return { id: String(r[0]), item: r[1], categoria: r[2], qtd: Number(r[3]) || 0, minimo: Number(r[4]) || 0, unidade: r[5], custo: Number(r[6]) || 0, situacao: r[7], atualizado: r[8] };
}

function lerFilamentos_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Filamentos');
  if (aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, ABAS.Filamentos.length).getValues()
    .filter(function (r) { return r[0] !== ''; }).map(filamentoDaLinha_);
}
function lerHistorico_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Histórico');
  if (aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, 7).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) { return { data: r[0] instanceof Date ? Utilities.formatDate(r[0], 'America/Sao_Paulo', 'yyyy-MM-dd') : String(r[0]), peca: r[1], material: r[2], cor: r[3], gramas: Number(r[4]) || 0, custo: Number(r[5]) || 0, descontou: r[6] }; });
}

function lerInsumos_() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Insumos');
  if (aba.getLastRow() < 2) return [];
  return aba.getRange(2, 1, aba.getLastRow() - 1, ABAS.Insumos.length).getValues()
    .filter(function (r) { return r[0] !== ''; }).map(insumoDaLinha_);
}

function acharLinha_(aba, id) {
  if (aba.getLastRow() < 2) return null;
  const ids = aba.getRange(2, 1, aba.getLastRow() - 1, 1).getValues();
  for (let k = 0; k < ids.length; k++) if (String(ids[k][0]) === String(id)) return k + 2;
  return null;
}

function novoId_() { return Utilities.getUuid().slice(0, 8); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
