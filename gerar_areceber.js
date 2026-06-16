// gerar_areceber.js — Cria abas Clientes, AReceber, FC_Recebimento e FC_Consolidado no Excel
// Node.js + exceljs + pdf_parser_comum
// Lê automaticamente o PDF "a receber.pdf" a cada execução

const ExcelJS = require('./node_modules/exceljs');
const { parsePDF, DATE_RE, CODE_RE, VALUE_RE } = require('./pdf_parser_comum');
const path = require('path');

const XLSX_PATH = path.join(__dirname, 'WP_Demonstrações_Financeiras.xlsx');
const PDF_PATH  = path.join(__dirname, 'a receber.pdf');

// ── Produtos reconhecidos no histórico ────────────────────────────────────────
const PRODUTOS = ['CONCRETO', 'AREIA', 'ARGAMASSA', 'CIMENTO'];
const PRODUTO_RE = new RegExp(`^(${PRODUTOS.join('|')})$`, 'i');

// Parser de linha específico para "a receber": além de doc/datas/cod/valor,
// tenta extrair o produto vendido pela posição (mesmo truque usado para datas/códigos/valores).
// Histórico real fica embaralhado nas linhas com múltiplos registros — extração é heurística.
function parseRowReceber(texts) {
  const dates    = texts.filter(t => DATE_RE.test(t));
  const codes    = texts.filter(t => CODE_RE.test(t));
  const values   = texts.filter(t => VALUE_RE.test(t));
  const produtos = texts.filter(t => PRODUTO_RE.test(t));

  if (!dates.length || !codes.length || !values.length) return [];

  const n = Math.floor(Math.min(dates.length / 2, codes.length, values.length));
  if (n < 1) return [];

  const firstDateIdx = texts.findIndex(t => DATE_RE.test(t));
  const docCandidates = texts.slice(0, firstDateIdx)
    .filter(t => !DATE_RE.test(t) && !CODE_RE.test(t) && !VALUE_RE.test(t));

  const half = Math.round(dates.length / 2);
  const emissaoDates = dates.slice(0, half);
  const vencDates    = dates.slice(half);

  const records = [];
  for (let i = 0; i < n; i++) {
    if (!vencDates[i] || !codes[i] || !values[i]) continue;
    const produto = produtos[i] ? produtos[i].toUpperCase() : 'Outros';
    records.push([
      docCandidates[i] || '',
      emissaoDates[i]  || '',
      vencDates[i],
      parseInt(codes[i]),
      '',       // nome preenchido depois pela tabela Clientes
      '',       // historico não extraído
      parseFloat(String(values[i]).replace(/\./g,'').replace(',','.')),
      produto.charAt(0) + produto.slice(1).toLowerCase()
    ]);
  }
  return records;
}

// ── Clientes conhecidos (nomes amigáveis) ─────────────────────────────────────
const CLIENTES = [
  [3780, 'MOISES ELIAS DA SILVA'],
  [4536, 'INCORPORACAO OPUS 208 SPE LTDA'],
  [4361, 'INCORPORACAO OPUS SPE LTDA'],
  [3940, 'CASAS DOURADO INCORPORADORA LTDA'],
  [4399, 'INCORPORACAO OPUS 72 SPE LTDA'],
];
const CLIENTES_MAP = {};
CLIENTES.forEach(([cod, nome]) => { CLIENTES_MAP[cod] = nome; });

function mesVenc(venc) {
  const [d,m,y] = venc.split('/');
  return `${y}-${m.padStart(2,'0')}`;
}

function mesLabel(yyyyMM) {
  const [y, m] = yyyyMM.split('-');
  const mNames = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  return `${mNames[parseInt(m)-1]}/${y.slice(2)}`;
}

async function main() {
  console.log(`📄 Lendo PDF: ${PDF_PATH}`);
  const DADOS = await parsePDF(PDF_PATH, parseRowReceber);

  // Preenche nome via Clientes (ou fallback)
  DADOS.forEach(r => { r[4] = CLIENTES_MAP[r[3]] || ('CLIENTE ' + r[3]); });

  // Meses dinâmicos a partir dos vencimentos reais
  const mesesSet = new Set(DADOS.map(r => mesVenc(r[2])));
  const MESES = Array.from(mesesSet).sort();
  const MESES_LABEL = MESES.map(mesLabel);

  console.log(`📊 Lendo Excel: ${XLSX_PATH}`);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);

  // ─── Remove abas existentes se houver ──────────────────────────────────────
  ['Clientes', 'AReceber', 'FC_Recebimento', 'FC_Consolidado'].forEach(name => {
    const s = wb.getWorksheet(name);
    if (s) wb.removeWorksheet(s.id);
  });

  // ════════════════════════════════════════════════════════════════════════════
  // ABA 1: Clientes
  // ════════════════════════════════════════════════════════════════════════════
  const wsCli = wb.addWorksheet('Clientes', { properties: { tabColor: { argb: 'FF548235' } } });
  wsCli.columns = [
    { header: 'Cod', key: 'cod', width: 10 },
    { header: 'Cliente', key: 'nome', width: 45 },
  ];
  ['A1','B1'].forEach(cell => {
    wsCli.getCell(cell).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF548235' } };
    wsCli.getCell(cell).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  });
  CLIENTES.forEach(([cod, nome]) => wsCli.addRow({ cod, nome }));
  const knownCods = new Set(CLIENTES.map(c => c[0]));
  const extraCods = new Set();
  DADOS.forEach(r => { if (!knownCods.has(r[3])) extraCods.add(r[3]); });
  extraCods.forEach(cod => {
    const found = DADOS.find(r => r[3] === cod);
    wsCli.addRow({ cod, nome: found[4] });
  });
  console.log(`✅ Clientes: ${wsCli.rowCount-1} clientes`);

  // ════════════════════════════════════════════════════════════════════════════
  // ABA 2: AReceber
  // ════════════════════════════════════════════════════════════════════════════
  const wsAR = wb.addWorksheet('AReceber', { properties: { tabColor: { argb: 'FF70AD47' } } });
  wsAR.columns = [
    { header: 'Documento', key: 'doc', width: 16 },
    { header: 'Emissao', key: 'emis', width: 13 },
    { header: 'Vencimento', key: 'venc', width: 13 },
    { header: 'MesVenc', key: 'mes', width: 10 },
    { header: 'Cliente_Cod', key: 'cod', width: 14 },
    { header: 'Cliente_Nome', key: 'nome', width: 40 },
    { header: 'Valor', key: 'valor', width: 16 },
    { header: 'Produto', key: 'produto', width: 16 },
    { header: 'Tipo', key: 'tipo', width: 12 },
    { header: 'GrupoMae', key: 'grupo', width: 22 },
  ];
  const hRowAR = wsAR.getRow(1);
  hRowAR.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF70AD47' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  });

  let totalRec = 0;
  DADOS.forEach(([doc, emis, venc, cod, nome, hist, valor, produto], i) => {
    const rowNum = i + 2;
    const mv = mesVenc(venc);
    totalRec += valor;
    const row = wsAR.addRow({
      doc: String(doc), emis, venc, mes: mv,
      cod, nome, valor, produto,
      tipo: 'Previsto', grupo: 'Entradas Operacionais'
    });
    row.getCell('valor').numFmt = '#,##0.00';
    if (rowNum % 2 === 0) {
      row.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
      });
    }
  });

  const totRowAR = wsAR.addRow({
    doc: 'TOTAL', emis: '', venc: '', mes: '',
    cod: '', nome: `${DADOS.length} registros`,
    valor: totalRec, produto: '', tipo: '', grupo: ''
  });
  totRowAR.font = { bold: true };
  totRowAR.getCell('valor').numFmt = '#,##0.00';
  totRowAR.getCell('valor').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } };

  console.log(`✅ AReceber: ${DADOS.length} registros | Total: R$ ${totalRec.toLocaleString('pt-BR',{minimumFractionDigits:2})}`);

  // ════════════════════════════════════════════════════════════════════════════
  // ABA 3: FC_Recebimento (pivot mês × produto)
  // ════════════════════════════════════════════════════════════════════════════
  const wsFR = wb.addWorksheet('FC_Recebimento', { properties: { tabColor: { argb: 'FF548235' } } });
  const pivotR = {}; // produto → mes → total
  DADOS.forEach(([, , venc, , , , valor, produto]) => {
    const mv = mesVenc(venc);
    if (!pivotR[produto]) pivotR[produto] = {};
    pivotR[produto][mv] = (pivotR[produto][mv] || 0) + valor;
  });

  const numFmt = '#,##0.00';
  const fcrHeaders = ['Produto', ...MESES_LABEL, 'TOTAL'];
  const hdrRowFR = wsFR.addRow(fcrHeaders);
  hdrRowFR.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF375623' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center' };
  });
  wsFR.getColumn(1).width = 22;
  MESES.forEach((_, i) => { wsFR.getColumn(i+2).width = 14; });
  wsFR.getColumn(MESES.length+2).width = 16;

  const produtosUsados = Object.keys(pivotR).sort();
  const recTotals = new Array(MESES.length).fill(0);
  produtosUsados.forEach(produto => {
    const row = [produto];
    let rowTotal = 0;
    MESES.forEach((m, mi) => {
      const v = (pivotR[produto] && pivotR[produto][m]) ? pivotR[produto][m] : 0;
      row.push(v);
      rowTotal += v;
      recTotals[mi] += v;
    });
    row.push(rowTotal);
    const r = wsFR.addRow(row);
    for (let c = 2; c <= MESES.length+2; c++) {
      r.getCell(c).numFmt = numFmt;
      r.getCell(c).alignment = { horizontal: 'right' };
    }
  });

  const grandRowFR = wsFR.addRow([
    '═══ TOTAL ENTRADAS', ...recTotals, recTotals.reduce((a,b)=>a+b,0)
  ]);
  grandRowFR.font = { bold: true, size: 12 };
  grandRowFR.eachCell((cell, colNumber) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF375623' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    if (colNumber > 1) cell.numFmt = numFmt;
  });

  console.log(`✅ FC_Recebimento gerado | Total: R$ ${recTotals.reduce((a,b)=>a+b,0).toLocaleString('pt-BR',{minimumFractionDigits:2})}`);

  // ════════════════════════════════════════════════════════════════════════════
  // ABA 4: FC_Consolidado (Entradas × Saídas × Saldo, lê FC_Direto já existente)
  // ════════════════════════════════════════════════════════════════════════════
  const wsFD = wb.getWorksheet('FC_Direto');
  let payTotalsByMonth = {}; // mesLabel → total saídas
  let payLabels = [];
  if (wsFD) {
    // Localiza a linha "═══ TOTAL SAÍDAS DE CAIXA" e o cabeçalho de meses
    const headerRow = wsFD.getRow(1);
    const headerLabels = [];
    headerRow.eachCell((cell, colNumber) => { headerLabels[colNumber] = String(cell.value || ''); });
    payLabels = headerLabels.filter(l => l && l !== 'Categoria' && l !== 'TOTAL');
    wsFD.eachRow((row) => {
      const firstCell = String(row.getCell(1).value || '');
      if (firstCell.includes('TOTAL SAÍDAS DE CAIXA')) {
        headerLabels.forEach((label, colNumber) => {
          if (colNumber > 1 && label !== 'TOTAL') {
            payTotalsByMonth[label] = row.getCell(colNumber).value || 0;
          }
        });
      }
    });
  } else {
    console.log('⚠️  Aba FC_Direto não encontrada — rode gerar_apagar.js antes para o Consolidado incluir as saídas.');
  }

  // União dos meses de AReceber (MESES_LABEL) e de FC_Direto (payLabels), ordenada cronologicamente
  const MES_NAMES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  function labelToSortKey(lbl) {
    const [mon, yy] = lbl.split('/');
    const mi = MES_NAMES.indexOf(mon);
    return `20${yy}-${String(mi+1).padStart(2,'0')}`;
  }
  const recByLabel = {};
  MESES_LABEL.forEach((lbl, i) => { recByLabel[lbl] = recTotals[i] || 0; });

  const unionLabels = Array.from(new Set([...MESES_LABEL, ...payLabels]))
    .sort((a,b) => labelToSortKey(a).localeCompare(labelToSortKey(b)));

  const wsCons = wb.addWorksheet('FC_Consolidado', { properties: { tabColor: { argb: 'FF1F4E79' } } });
  const consHeaders = ['Indicador', ...unionLabels, 'TOTAL'];
  const hdrRowCons = wsCons.addRow(consHeaders);
  hdrRowCons.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center' };
  });
  wsCons.getColumn(1).width = 28;
  unionLabels.forEach((_, i) => { wsCons.getColumn(i+2).width = 14; });
  wsCons.getColumn(unionLabels.length+2).width = 16;

  const entradasArr = unionLabels.map(lbl => recByLabel[lbl] || 0);
  const saidasArr   = unionLabels.map(lbl => payTotalsByMonth[lbl] || 0);
  const saldoArr    = unionLabels.map((lbl, i) => entradasArr[i] - saidasArr[i]);
  let acumulado = 0;
  const acumArr = saldoArr.map(v => (acumulado += v));

  function addConsRow(label, values, opts={}) {
    const row = wsCons.addRow([label, ...values, values.reduce((a,b)=>a+b,0)]);
    for (let c = 2; c <= unionLabels.length+2; c++) row.getCell(c).numFmt = numFmt;
    if (opts.bold) row.font = { bold: true };
    if (opts.fill) row.eachCell((cell, colNumber) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.fill } };
      if (opts.fontColor) cell.font = { bold: true, color: { argb: opts.fontColor } };
    });
    return row;
  }

  addConsRow('Entradas (a receber)', entradasArr, { fill: 'FFE2EFDA' });
  addConsRow('Saídas (a pagar)', saidasArr, { fill: 'FFFCE4D6' });
  addConsRow('= Saldo do Mês', saldoArr, { bold: true, fill: 'FFD9E1F2' });
  const accRow = wsCons.addRow(['Saldo Acumulado', ...acumArr, acumArr[acumArr.length-1] || 0]);
  accRow.font = { bold: true, size: 12 };
  accRow.eachCell((cell, colNumber) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    if (colNumber > 1) cell.numFmt = numFmt;
  });

  const saldoTotal = entradasArr.reduce((a,b)=>a+b,0) - saidasArr.reduce((a,b)=>a+b,0);
  console.log(`✅ FC_Consolidado gerado | Saldo projetado: R$ ${saldoTotal.toLocaleString('pt-BR',{minimumFractionDigits:2})}`);

  // Save
  await wb.xlsx.writeFile(XLSX_PATH);
  console.log(`\n🎉 Excel salvo com sucesso: ${XLSX_PATH}`);
  console.log(`   Abas adicionadas: Clientes | AReceber | FC_Recebimento | FC_Consolidado`);
}

main().catch(err => { console.error('❌ Erro:', err.message); process.exit(1); });
