// gerar_recebidos.js — Cria abas Recebidos, FC_Recebimentos, FC_Realizado e FC_Comparativo
// Node.js + exceljs + pdf_parser_comum
// Lê todos os PDFs da pasta pdfs/recebidos/
// EXECUTAR APÓS gerar_pagos.js (FC_Realizado precisa de FC_Pagos)

const ExcelJS = require('./node_modules/exceljs');
const { parsePDF, DATE_RE, CODE_RE, VALUE_RE } = require('./pdf_parser_comum');
const path = require('path');
const fs   = require('fs');

const XLSX_PATH = path.join(__dirname, 'WP_Demonstrações_Financeiras.xlsx');
const PDF_DIR   = path.join(__dirname, 'pdfs', 'recebidos');

// ── Produtos reconhecidos no histórico ────────────────────────────────────────
const PRODUTOS = ['CONCRETO', 'AREIA', 'ARGAMASSA', 'CIMENTO'];
const PRODUTO_RE = new RegExp(`^(${PRODUTOS.join('|')})$`, 'i');
const PRODUTO_REMAP = { 'CIMENTO': 'AREIA' };

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
    let produto = produtos[i] ? produtos[i].toUpperCase() : 'OUTROS';
    produto = PRODUTO_REMAP[produto] || produto;
    records.push([
      docCandidates[i] || '',
      emissaoDates[i]  || '',
      vencDates[i],
      parseInt(codes[i]),
      '',
      '',
      parseFloat(String(values[i]).replace(/\./g,'').replace(',','.')),
      produto.charAt(0) + produto.slice(1).toLowerCase()
    ]);
  }
  return records;
}

// ── Clientes conhecidos ───────────────────────────────────────────────────────
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

const MES_NAMES = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
function labelToSortKey(lbl) {
  const [mon, yy] = lbl.split('/');
  const mi = MES_NAMES.indexOf(mon);
  return `20${yy}-${String(mi+1).padStart(2,'0')}`;
}

// Lê linha "═══ TOTAL ..." de uma aba pivot e retorna {label → valor}
function readTotalRow(ws, keyword) {
  const headerRow  = ws.getRow(1);
  const headerLabels = [];
  headerRow.eachCell((cell, colNumber) => { headerLabels[colNumber] = String(cell.value || ''); });
  const payLabels = headerLabels.filter(l => l && l !== 'Categoria' && l !== 'Produto' && l !== 'Indicador' && l !== 'TOTAL');

  let result = {};
  ws.eachRow(row => {
    const firstCell = String(row.getCell(1).value || '');
    if (firstCell.includes(keyword)) {
      headerLabels.forEach((label, colNumber) => {
        if (colNumber > 1 && label !== 'TOTAL') {
          result[label] = parseFloat(row.getCell(colNumber).value) || 0;
        }
      });
    }
  });
  return { byLabel: result, labels: payLabels };
}

async function main() {
  if (!fs.existsSync(PDF_DIR)) {
    console.error(`❌ Pasta não encontrada: ${PDF_DIR}`);
    process.exit(1);
  }
  const pdfFiles = fs.readdirSync(PDF_DIR)
    .filter(f => f.toLowerCase().endsWith('.pdf'))
    .map(f => path.join(PDF_DIR, f));

  if (!pdfFiles.length) {
    console.error(`❌ Nenhum PDF encontrado em: ${PDF_DIR}`);
    process.exit(1);
  }

  let allRecords = [];
  for (const pdfFile of pdfFiles) {
    console.log(`📄 Lendo: ${path.basename(pdfFile)}`);
    const records = await parsePDF(pdfFile, parseRowReceber);
    allRecords = allRecords.concat(records);
  }

  const seenKeys = new Set();
  const DADOS = allRecords.filter(r => {
    const k = `${r[0]}|${r[3]}|${r[6]}`;
    if (seenKeys.has(k)) return false;
    seenKeys.add(k); return true;
  });
  console.log(`✅ ${DADOS.length} registros únicos de ${pdfFiles.length} arquivo(s) (${allRecords.length - DADOS.length} duplicatas removidas)`);

  DADOS.forEach(r => { r[4] = CLIENTES_MAP[r[3]] || ('CLIENTE ' + r[3]); });

  const mesesSet = new Set(DADOS.map(r => mesVenc(r[2])));
  const MESES = Array.from(mesesSet).sort();
  const MESES_LABEL = MESES.map(mesLabel);

  console.log(`📊 Lendo Excel: ${XLSX_PATH}`);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);

  ['Recebidos','FC_Recebimentos','FC_Realizado','FC_Comparativo'].forEach(name => {
    const s = wb.getWorksheet(name);
    if (s) wb.removeWorksheet(s.id);
  });

  // ════════════════════════════════════════════════════════════════════════════
  // ABA 1: Recebidos
  // ════════════════════════════════════════════════════════════════════════════
  const wsR = wb.addWorksheet('Recebidos', { properties: { tabColor: { argb: 'FF1A5276' } } });
  wsR.columns = [
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
  wsR.getRow(1).eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A5276' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  });

  let totalRec = 0;
  DADOS.forEach(([doc, emis, venc, cod, nome, , valor, produto], i) => {
    const mv = mesVenc(venc);
    totalRec += valor;
    const row = wsR.addRow({ doc: String(doc), emis, venc, mes: mv, cod, nome, valor, produto, tipo: 'Realizado', grupo: 'Entradas Operacionais' });
    row.getCell('valor').numFmt = '#,##0.00';
    if ((i+2) % 2 === 0) {
      row.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD6EAF8' } };
      });
    }
  });

  const totRowR = wsR.addRow({ doc: 'TOTAL', nome: `${DADOS.length} registros`, valor: totalRec });
  totRowR.font = { bold: true };
  totRowR.getCell('valor').numFmt = '#,##0.00';
  totRowR.getCell('valor').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } };
  console.log(`✅ Recebidos: ${DADOS.length} registros | Total: R$ ${totalRec.toLocaleString('pt-BR',{minimumFractionDigits:2})}`);

  // ════════════════════════════════════════════════════════════════════════════
  // ABA 2: FC_Recebimentos (pivot mês × produto)
  // ════════════════════════════════════════════════════════════════════════════
  const wsFR = wb.addWorksheet('FC_Recebimentos', { properties: { tabColor: { argb: 'FF145A32' } } });
  const pivotR = {};
  DADOS.forEach(([,,venc,,,, valor, produto]) => {
    const mv = mesVenc(venc);
    if (!pivotR[produto]) pivotR[produto] = {};
    pivotR[produto][mv] = (pivotR[produto][mv] || 0) + valor;
  });

  const numFmt = '#,##0.00';
  const hdrRowFR = wsFR.addRow(['Produto', ...MESES_LABEL, 'TOTAL']);
  hdrRowFR.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF145A32' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center' };
  });
  wsFR.getColumn(1).width = 22;
  MESES.forEach((_, i) => { wsFR.getColumn(i+2).width = 14; });
  wsFR.getColumn(MESES.length+2).width = 16;

  const produtosUsados = Object.keys(pivotR).sort();
  const recTotals = new Array(MESES.length).fill(0);
  produtosUsados.forEach(produto => {
    const row = [produto]; let rowTotal = 0;
    MESES.forEach((m, mi) => {
      const v = (pivotR[produto] && pivotR[produto][m]) ? pivotR[produto][m] : 0;
      row.push(v); rowTotal += v; recTotals[mi] += v;
    });
    row.push(rowTotal);
    const r = wsFR.addRow(row);
    for (let c = 2; c <= MESES.length+2; c++) {
      r.getCell(c).numFmt = numFmt;
      r.getCell(c).alignment = { horizontal: 'right' };
    }
  });

  const grandRowFR = wsFR.addRow(['═══ TOTAL ENTRADAS', ...recTotals, recTotals.reduce((a,b)=>a+b,0)]);
  grandRowFR.eachCell((cell, colNumber) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF145A32' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    if (colNumber > 1) cell.numFmt = numFmt;
  });
  console.log(`✅ FC_Recebimentos gerado | Total: R$ ${recTotals.reduce((a,b)=>a+b,0).toLocaleString('pt-BR',{minimumFractionDigits:2})}`);

  // ════════════════════════════════════════════════════════════════════════════
  // ABA 3: FC_Realizado (Consolidado Realizado: Recebidos × Pagos)
  // ════════════════════════════════════════════════════════════════════════════
  const recByLabel = {};
  MESES_LABEL.forEach((lbl, i) => { recByLabel[lbl] = recTotals[i] || 0; });

  const wsFPagos = wb.getWorksheet('FC_Pagos');
  let pagTotalsByMonth = {};
  let pagLabels = [];
  if (wsFPagos) {
    const r = readTotalRow(wsFPagos, 'TOTAL SAÍDAS DE CAIXA');
    pagTotalsByMonth = r.byLabel;
    pagLabels = r.labels;
  } else {
    console.log('⚠️  Aba FC_Pagos não encontrada — execute gerar_pagos.js antes deste script.');
  }

  const unionLabelsReal = Array.from(new Set([...MESES_LABEL, ...pagLabels]))
    .sort((a,b) => labelToSortKey(a).localeCompare(labelToSortKey(b)));

  const wsCons = wb.addWorksheet('FC_Realizado', { properties: { tabColor: { argb: 'FF7D6608' } } });
  const consHeaders = ['Indicador', ...unionLabelsReal, 'TOTAL'];
  const hdrRowCons = wsCons.addRow(consHeaders);
  hdrRowCons.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7D6608' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center' };
  });
  wsCons.getColumn(1).width = 28;
  unionLabelsReal.forEach((_, i) => { wsCons.getColumn(i+2).width = 14; });
  wsCons.getColumn(unionLabelsReal.length+2).width = 16;

  const entradasReal = unionLabelsReal.map(lbl => recByLabel[lbl] || 0);
  const saidasReal   = unionLabelsReal.map(lbl => pagTotalsByMonth[lbl] || 0);
  const saldoReal    = unionLabelsReal.map((_,i) => entradasReal[i] - saidasReal[i]);
  let acum = 0;
  const acumReal = saldoReal.map(v => (acum += v));

  function addConsRow(ws, label, values, opts={}) {
    const row = ws.addRow([label, ...values, values.reduce((a,b)=>a+b,0)]);
    for (let c = 2; c <= values.length+2; c++) row.getCell(c).numFmt = numFmt;
    if (opts.bold) row.font = { bold: true };
    if (opts.fill) row.eachCell((cell, colNumber) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.fill } };
      if (opts.fontColor) cell.font = { bold: true, color: { argb: opts.fontColor } };
    });
  }

  addConsRow(wsCons, 'Entradas (recebido)', entradasReal, { fill: 'FFD6EAF8' });
  addConsRow(wsCons, 'Saídas (pago)', saidasReal, { fill: 'FFFCE4D6' });
  addConsRow(wsCons, '= Saldo do Mês', saldoReal, { bold: true, fill: 'FFD9E1F2' });
  const accRow = wsCons.addRow(['Saldo Acumulado', ...acumReal, acumReal[acumReal.length-1] || 0]);
  accRow.eachCell((cell, colNumber) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7D6608' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    if (colNumber > 1) cell.numFmt = numFmt;
  });

  const saldoTotalReal = entradasReal.reduce((a,b)=>a+b,0) - saidasReal.reduce((a,b)=>a+b,0);
  console.log(`✅ FC_Realizado gerado | Saldo: R$ ${saldoTotalReal.toLocaleString('pt-BR',{minimumFractionDigits:2})}`);

  // ════════════════════════════════════════════════════════════════════════════
  // ABA 4: FC_Comparativo (Orçado × Realizado)
  // ════════════════════════════════════════════════════════════════════════════
  // Lê totais orçados das abas FC_Recebimento (entradas) e FC_Direto (saídas)
  let orcEntradasByLabel = {};
  let orcEntradasLabels  = [];
  const wsFRecebimento = wb.getWorksheet('FC_Recebimento');
  if (wsFRecebimento) {
    const r = readTotalRow(wsFRecebimento, 'TOTAL ENTRADAS');
    orcEntradasByLabel = r.byLabel;
    orcEntradasLabels  = r.labels;
  }

  let orcSaidasByLabel = {};
  let orcSaidasLabels  = [];
  const wsFDireto = wb.getWorksheet('FC_Direto');
  if (wsFDireto) {
    const r = readTotalRow(wsFDireto, 'TOTAL SAÍDAS DE CAIXA');
    orcSaidasByLabel = r.byLabel;
    orcSaidasLabels  = r.labels;
  }

  // União de todos os meses de todos os datasets
  const allLabels = Array.from(new Set([
    ...orcEntradasLabels, ...orcSaidasLabels,
    ...MESES_LABEL, ...pagLabels
  ])).sort((a,b) => labelToSortKey(a).localeCompare(labelToSortKey(b)));

  const wsComp = wb.addWorksheet('FC_Comparativo', { properties: { tabColor: { argb: 'FF4A235A' } } });
  const compHeaders = ['Indicador', ...allLabels, 'TOTAL'];
  const hdrComp = wsComp.addRow(compHeaders);
  hdrComp.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4A235A' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center' };
  });
  wsComp.getColumn(1).width = 30;
  allLabels.forEach((_, i) => { wsComp.getColumn(i+2).width = 14; });
  wsComp.getColumn(allLabels.length+2).width = 16;

  // Arrays por mês para cada linha
  const orcEntArr  = allLabels.map(l => orcEntradasByLabel[l] || 0);
  const realEntArr = allLabels.map(l => recByLabel[l] || 0);
  const deltaEnt   = allLabels.map((_,i) => realEntArr[i] - orcEntArr[i]);

  const orcSaiArr  = allLabels.map(l => orcSaidasByLabel[l] || 0);
  const realSaiArr = allLabels.map(l => pagTotalsByMonth[l] || 0);
  const deltaSai   = allLabels.map((_,i) => realSaiArr[i] - orcSaiArr[i]);

  const orcSaldoArr  = allLabels.map((_,i) => orcEntArr[i] - orcSaiArr[i]);
  const realSaldoArr = allLabels.map((_,i) => realEntArr[i] - realSaiArr[i]);
  const deltaSaldo   = allLabels.map((_,i) => realSaldoArr[i] - orcSaldoArr[i]);

  function addCompRow(label, values, opts={}) {
    const row = wsComp.addRow([label, ...values, values.reduce((a,b)=>a+b,0)]);
    for (let c = 2; c <= values.length+2; c++) row.getCell(c).numFmt = numFmt;
    if (opts.bold) row.font = { bold: true };
    if (opts.fill) row.eachCell((cell, colNumber) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.fill } };
      if (opts.fontColor) cell.font = { bold: true, color: { argb: opts.fontColor } };
    });
  }

  // Seção Entradas
  const secEnt = wsComp.addRow(['── ENTRADAS']);
  secEnt.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF145A32' } };
  secEnt.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

  addCompRow('Entradas Orçado',    orcEntArr,  { fill: 'FFE2EFDA' });
  addCompRow('Entradas Realizado', realEntArr, { fill: 'FFD6EAF8' });
  addCompRow('Δ Entradas (Real − Orç)', deltaEnt, { bold: true, fill: 'FFF0F0F0' });

  wsComp.addRow([]);

  // Seção Saídas
  const secSai = wsComp.addRow(['── SAÍDAS']);
  secSai.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9F1239' } };
  secSai.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

  addCompRow('Saídas Orçado',    orcSaiArr,  { fill: 'FFFCE4D6' });
  addCompRow('Saídas Realizado', realSaiArr, { fill: 'FFFDE8D8' });
  addCompRow('Δ Saídas (Real − Orç)', deltaSai, { bold: true, fill: 'FFF0F0F0' });

  wsComp.addRow([]);

  // Seção Saldo
  const secSaldo = wsComp.addRow(['── SALDO']);
  secSaldo.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
  secSaldo.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

  addCompRow('Saldo Orçado',    orcSaldoArr,  { fill: 'FFD9E1F2' });
  addCompRow('Saldo Realizado', realSaldoArr, { fill: 'FFD6EAF8' });
  addCompRow('Δ Saldo (Real − Orç)', deltaSaldo, { bold: true, fill: 'FFF0F0F0' });

  console.log(`✅ FC_Comparativo gerado`);

  await wb.xlsx.writeFile(XLSX_PATH);
  console.log(`\n🎉 Excel salvo com sucesso: ${XLSX_PATH}`);
  console.log(`   Abas adicionadas: Recebidos | FC_Recebimentos | FC_Realizado | FC_Comparativo`);
}

main().catch(err => { console.error('❌ Erro:', err.message); process.exit(1); });
