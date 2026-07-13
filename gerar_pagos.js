// gerar_pagos.js — Cria abas Pagos e FC_Pagos no Excel (Fluxo Realizado - Saídas)
// Node.js + exceljs + pdf_parser_comum
// Lê todos os PDFs da pasta pdfs/pagos/

const ExcelJS = require('./node_modules/exceljs');
const { parsePDF, DATE_RE, CODE_RE, VALUE_RE } = require('./pdf_parser_comum');

// Parser específico para "Relatório de Contas Pagas" do Gdoor
// Colunas: Documento | Emissão | Vencimento | Pagamento | Fornecedor | Portador
//        | Valor da parcela | Juros | Descontos | Valor Pago
// → 3 datas por registro, 4 valores por registro. Queremos "Valor Pago" (último grupo).
function parseRowPagos(texts) {
  const dates  = texts.filter(t => DATE_RE.test(t));
  const codes  = texts.filter(t => CODE_RE.test(t));
  const values = texts.filter(t => VALUE_RE.test(t));

  if (!dates.length || !codes.length || !values.length) return [];

  // n = número de registros na linha
  const n = Math.floor(Math.min(dates.length / 3, codes.length, values.length / 4));
  if (n < 1) return [];

  const firstDateIdx = texts.findIndex(t => DATE_RE.test(t));
  const docCandidates = texts.slice(0, firstDateIdx)
    .filter(t => !DATE_RE.test(t) && !CODE_RE.test(t) && !VALUE_RE.test(t));

  // datas: [emissao x n] [vencimento x n] [pagamento x n]
  const emissaoDates = dates.slice(0, n);
  const pagDates     = dates.slice(2 * n, 3 * n);

  // valores: últimos n = Valor Pago
  const valorPagoStart = values.length - n;

  // Tenta extrair nome do fornecedor do PDF: tokens entre a última data e o primeiro código
  let lastDateIdx = -1;
  texts.forEach((t, i) => { if (DATE_RE.test(t)) lastDateIdx = i; });
  const firstCodeIdx = texts.findIndex(t => CODE_RE.test(t));
  const midTokens = texts
    .slice(lastDateIdx + 1, firstCodeIdx >= 0 ? firstCodeIdx : texts.length)
    .filter(t => !VALUE_RE.test(t) && t.trim().length > 1);
  // midTokens contém: [nomeForns..., portadores...]. Pega os primeiros (antes dos portadores).
  // Heurística: portador é curto (banco/espécie); nome de empresa tende a ser maior.
  // Divide midTokens em 2*n partes e pega a primeira metade como nome.
  const halfLen = Math.ceil(midTokens.length / 2);
  const nameTokensAll = midTokens.slice(0, halfLen);
  const chunkSize = Math.max(1, Math.ceil(nameTokensAll.length / n));

  const records = [];
  for (let i = 0; i < n; i++) {
    const pagDate   = pagDates[i];
    const valorPago = values[valorPagoStart + i];
    if (!pagDate || !codes[i] || !valorPago) continue;
    const pdfName = nameTokensAll.slice(i * chunkSize, (i + 1) * chunkSize).join(' ');
    records.push([
      docCandidates[i] || '',
      emissaoDates[i]  || '',
      pagDate,
      parseInt(codes[i]),
      pdfName,   // nome extraído do PDF (fallback se não estiver em DEPARA_NOME)
      '',
      parseFloat(String(valorPago).replace(/\./g,'').replace(',','.'))
    ]);
  }
  return records;
}
const path = require('path');
const fs   = require('fs');

const XLSX_PATH = path.join(__dirname, 'WP_Demonstrações_Financeiras.xlsx');
const PDF_DIR   = path.join(__dirname, 'pdfs', 'pagos');

// ── DE-PARA: mesmo mapeamento de gerar_apagar.js ─────────────────────────────
const DEPARA = [
  [3478, 'VOTORANTIM CIMENTOS S/A',               'Matérias-Primas'],
  [3481, 'PEDRA BRITADA IND COM LTDA',             'Matérias-Primas'],
  [3768, 'PEDREIRA CAMPO LIMPO LTDA',              'Matérias-Primas'],
  [3953, 'PEDREIRA ANAPOLIS LTDA',                 'Matérias-Primas'],
  [3938, 'MF MINERACAO E TRANSPORTE LTDA',         'Matérias-Primas'],
  [3855, 'M R TRANSPORTES E CONSULTORIA LTDA',     'Matérias-Primas'],
  [3514, 'MATCHEM SP PRODUTOS QUIMICOS LTDA',      'Matérias-Primas'],
  [3497, 'ERCA IND E COM DE PRODUTOS QUIMICO',     'Matérias-Primas'],
  [3407, 'JL QUIMICA IND E COM LTDA',              'Matérias-Primas'],
  [3951, 'FARMACIA DA CONSTRUCAO LTDA',            'Matérias-Primas'],
  [3719, 'IMPER KING IMPERMEABILIZANTES LTDA',     'Matérias-Primas'],
  [3502, 'IPIRANGA PRODUTOS DE PETROLEO S/A',      'Combustível'],
  [3694, 'ABER LOGISTICA LTDA',                    'Fretes e Logística'],
  [3683, 'LINO BERTOLDO TRANSPORTES LTDA',         'Fretes e Logística'],
  [2934, 'BANCO VOLVO (BRASIL) S/A',               'Empréstimos / Financiamentos'],
  [3526, 'PUTZMEISTER BRASIL LTDA',                'Empréstimos / Financiamentos'],
  [2864, 'FGTS - FUNDO DE GARANTIA DO TEMPO DE SERVICO', 'Impostos e Encargos'],
  [2846, 'INSTITUTO NACIONAL DO SEGURO SOCIAL',    'Impostos e Encargos'],
  [3290, 'PREFEITURA DE GOIANIA - SECRETARIA D',   'Impostos e Encargos'],
  [3494, 'INST BRASIL MEIO AMBIENTE (IBAMA)',       'Impostos e Encargos'],
  [3838, 'NOTA CONTROL TECNOLOGIA LTDA',           'Impostos e Encargos'],
  [3745, 'ESSOR SEGUROS S.A.',                     'Seguros'],
  [3373, 'SEM PARAR INSTITUICAO DE PAGAMENTO',     'Pedágio'],
  [3440, 'EQUATORIAL GOIAS DISTRIBUIDORA DE ENERGIA','Serviços'],
  [3360, 'TURBOFI TELECOM LTDA (15951)',            'Serviços'],
  [3753, 'TURBOFI TELECOM LTDA (22394)',            'Serviços'],
  [3519, 'UNIODONTO GOIANIA COOP DE CIRURGIO',     'Serviços'],
  [3413, 'TRIO CONSULTORIA TECNICA DE ENGENH',     'Serviços'],
  [3821, 'CONCRE-GOIAS LOCACOES LTDA',             'Serviços'],
  [3381, 'J.S GOIS TACOGRAFOS EIRELI',             'Serviços'],
  [3761, 'MARCELO NUNES DE OLIVEIRA LTDA',         'Serviços'],
  [3600, 'CDP SERV DE MOTORES E BOMBAS LTDA',      'Serviços'],
  [3422, 'AUTO ELETRICA LILI LTDA',                'Serviços'],
  [3408, 'MINORU MARIO YAMADA EIRELI ME',          'Serviços'],
  [3311, 'TEK FABRIL EIRELLI EPP',                 'Serviços'],
  [3863, 'STAR SHOP ATACADAO DA LIMPEZA LTDA',     'Serviços'],
  [3900, 'BELA COMERCIO IMPORTACAO E EXPORTA',     'Serviços'],
  [3917, 'VITORIA SERVICOS HIDRAULICOS LTDA',      'Serviços'],
  [3645, 'CENTRAL PECAS E ACESSORIOS PARA CAM',    'Manutenção / Peças'],
  [3393, 'SCAMPEÇAS LTDA',                         'Manutenção / Peças'],
  [3348, 'AGUILERA AUTO PEÇAS DE GOIAS LTDA',      'Manutenção / Peças'],
  [3300, 'CARRETEIRO AUTO CENTER LTDA',            'Manutenção / Peças'],
  [2838, 'SUECIA VEICULOS S.A.',                   'Manutenção / Peças'],
  [2823, 'BELCAR CAMINHOES E MAQUINAS LTDA',       'Manutenção / Peças'],
  [3939, 'TRACTOR PECAS E SOLUCOES PARA MAQU',    'Manutenção / Peças'],
  [3379, 'TRACTOR LIDER DISTR PCAS P/TRATORES',   'Manutenção / Peças'],
  [2987, 'REGIONAL DIST MOLAS LTDA',               'Manutenção / Peças'],
  [3876, 'MIX PECAS E SERVICOS DE BOMBAS PARA',   'Manutenção / Peças'],
  [3500, 'BOMBEAR CENTRO OESTE LTDA',              'Manutenção / Peças'],
  [2894, 'EXPRESS REFORMA E COMERCIO DE PNEU',     'Manutenção / Peças'],
  [3561, 'BIGA ACESSORIOS RACUDAO LTDA',           'Manutenção / Peças'],
  [3354, 'DRJ COM FILTROS E PCAS EIRELI',          'Manutenção / Peças'],
  [3293, 'DCCO SOLUCOES EM ENERGIA E EQUIPAM',     'Manutenção / Peças'],
  [3707, 'RODOGYN IMPLEMENTOS RODOVIARIOS L',      'Manutenção / Peças'],
  [2801, 'REDEMIL IMPLEMENTOS RODOVIARIOS LT',     'Manutenção / Peças'],
  [3565, 'REINO DA BORRACHA LTDA',                 'Manutenção / Peças'],
  [3351, 'MARK OLEO LUBRIFICANTE EIRELLI ME',      'Manutenção / Peças'],
  [3728, 'ICONIC LUBRIFICANTES S.A.',              'Manutenção / Peças'],
  [2916, 'CENTRO OESTE EMBREAGENS LTDA',           'Manutenção / Peças'],
  [3207, 'BANDEIRA PECAS E SERVICOS LTDA',         'Manutenção / Peças'],
  [3435, 'MOREIRA COM IMP E EXP LTDA ME',         'Manutenção / Peças'],
  [3702, 'PENTAGONO EQUIPAMENTOS DE PROTECA',      'Manutenção / Peças'],
  [3396, 'ALR ELETRICA LTDA',                      'Manutenção / Peças'],
  [3944, 'VEDACIL COMPONENTES HIDRAULICOS LT',     'Manutenção / Peças'],
  [3366, 'VEDACIL COMPONENTES HIDRAULICOS LT',     'Manutenção / Peças'],
  [3956, 'CASA DOS POSTOS COMERCIO E SERVICO',     'Manutenção / Peças'],
  [3390, 'CARVALHO COM DE TINTAS LTDA ME',         'Outros Operacionais'],
  [3492, 'JVC DISTR EIRELI EPP',                   'Outros Operacionais'],
  [3398, 'BATERIAS PERIM LTDA',                    'Manutenção / Peças'],
];

const DEPARA_MAP  = {};
const DEPARA_NOME = {};
DEPARA.forEach(([cod, nome, cat]) => { DEPARA_MAP[cod] = cat; DEPARA_NOME[cod] = nome; });

function getCategoria(cod) { return DEPARA_MAP[cod] || 'Outros Operacionais'; }
function getGrupoMae(cat) {
  return cat === 'Empréstimos / Financiamentos' ? 'Empréstimos / Financiamentos' : 'Saídas Operacionais';
}

const CATS_OPERACIONAL = [
  'Matérias-Primas','Combustível','Fretes e Logística','Manutenção / Peças',
  'Serviços','Seguros','Pedágio','Impostos e Encargos','Outros Operacionais',
];

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
    const records = await parsePDF(pdfFile, parseRowPagos);
    allRecords = allRecords.concat(records);
  }

  const seenKeys = new Set();
  const DADOS = allRecords.filter(r => {
    const k = `${r[0]}|${r[3]}|${r[6]}`;
    if (seenKeys.has(k)) return false;
    seenKeys.add(k); return true;
  });
  console.log(`✅ ${DADOS.length} registros únicos de ${pdfFiles.length} arquivo(s) (${allRecords.length - DADOS.length} duplicatas removidas)`);

  const mesesSet = new Set(DADOS.map(r => mesVenc(r[2])));
  const MESES = Array.from(mesesSet).sort();
  const MESES_LABEL = MESES.map(mesLabel);

  console.log(`📊 Lendo Excel: ${XLSX_PATH}`);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_PATH);

  // Complementa DEPARA_NOME com nomes da aba DePara (criada pelo gerar_apagar.js)
  const wsDePara = wb.getWorksheet('DePara');
  if (wsDePara) {
    wsDePara.eachRow((row, rowNum) => {
      if (rowNum === 1) return;
      const cod  = parseInt(row.getCell(1).value);
      const nome = String(row.getCell(2).value || '').trim();
      if (cod && nome && !DEPARA_NOME[cod]) DEPARA_NOME[cod] = nome;
    });
  }
  DADOS.forEach(r => { r[4] = DEPARA_NOME[r[3]] || r[4] || ('FORNECEDOR ' + r[3]); });

  ['Pagos','FC_Pagos'].forEach(name => {
    const s = wb.getWorksheet(name);
    if (s) wb.removeWorksheet(s.id);
  });

  // ════════════════════════════════════════════════════════════════════════════
  // ABA 1: Pagos
  // ════════════════════════════════════════════════════════════════════════════
  const wsP = wb.addWorksheet('Pagos', { properties: { tabColor: { argb: 'FF9F1239' } } });
  wsP.columns = [
    { header: 'Documento', key: 'doc', width: 16 },
    { header: 'Emissao', key: 'emis', width: 13 },
    { header: 'Vencimento', key: 'venc', width: 13 },
    { header: 'MesVenc', key: 'mes', width: 10 },
    { header: 'Fornecedor_Cod', key: 'cod', width: 16 },
    { header: 'Fornecedor_Nome', key: 'nome', width: 45 },
    { header: 'Historico', key: 'hist', width: 40 },
    { header: 'Valor', key: 'valor', width: 16 },
    { header: 'Categoria', key: 'cat', width: 30 },
    { header: 'Tipo', key: 'tipo', width: 12 },
    { header: 'GrupoMae', key: 'grupo', width: 30 },
  ];
  const hRowP = wsP.getRow(1);
  hRowP.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9F1239' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  });

  let totalVal = 0;
  DADOS.forEach(([doc, emis, venc, cod, nome, hist, valor], i) => {
    const cat   = getCategoria(cod);
    const grupo = getGrupoMae(cat);
    const mv    = mesVenc(venc);
    totalVal += valor;
    const row = wsP.addRow({ doc: String(doc), emis, venc, mes: mv, cod, nome, hist, valor, cat, tipo: 'Realizado', grupo });
    row.getCell('valor').numFmt = '#,##0.00';
    if ((i+2) % 2 === 0) {
      row.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } };
      });
    }
  });

  const totRowP = wsP.addRow({ doc: 'TOTAL', nome: `${DADOS.length} registros`, valor: totalVal });
  totRowP.font = { bold: true };
  totRowP.getCell('valor').numFmt = '#,##0.00';
  totRowP.getCell('valor').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
  console.log(`✅ Pagos: ${DADOS.length} registros | Total: R$ ${totalVal.toLocaleString('pt-BR',{minimumFractionDigits:2})}`);

  // ════════════════════════════════════════════════════════════════════════════
  // ABA 2: FC_Pagos (pivot mês × categoria)
  // ════════════════════════════════════════════════════════════════════════════
  const wsFC = wb.addWorksheet('FC_Pagos', { properties: { tabColor: { argb: 'FF7F0000' } } });

  const pivot = {};
  DADOS.forEach(([,,venc,cod,,,valor]) => {
    const cat = getCategoria(cod);
    const mv  = mesVenc(venc);
    if (!pivot[cat]) pivot[cat] = {};
    pivot[cat][mv] = (pivot[cat][mv] || 0) + valor;
  });

  const fcHeaders = ['Categoria', ...MESES_LABEL, 'TOTAL'];
  const hdrRow = wsFC.addRow(fcHeaders);
  hdrRow.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center' };
  });
  wsFC.getColumn(1).width = 35;
  MESES.forEach((_, i) => { wsFC.getColumn(i+2).width = 14; });
  wsFC.getColumn(MESES.length+2).width = 16;

  const numFmt = '#,##0.00';

  function addSection(label, cats, bgColor) {
    const sHdr = wsFC.addRow([label]);
    sHdr.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    sHdr.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

    const sectionTotals = new Array(MESES.length).fill(0);
    cats.forEach(cat => {
      const row = [cat]; let rowTotal = 0;
      MESES.forEach((m, mi) => {
        const v = (pivot[cat] && pivot[cat][m]) ? pivot[cat][m] : 0;
        row.push(v); rowTotal += v; sectionTotals[mi] += v;
      });
      row.push(rowTotal);
      const r = wsFC.addRow(row);
      r.getCell(1).alignment = { indent: 2 };
      for (let c = 2; c <= MESES.length+2; c++) {
        r.getCell(c).numFmt = numFmt;
        r.getCell(c).alignment = { horizontal: 'right' };
      }
    });
    const totals = ['= Total ' + label.replace(/^[^\w]*/, ''), ...sectionTotals, sectionTotals.reduce((a,b)=>a+b,0)];
    const tRow = wsFC.addRow(totals);
    tRow.font = { bold: true };
    tRow.eachCell((cell, colNumber) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
      if (colNumber > 1) cell.numFmt = numFmt;
    });
    wsFC.addRow([]);
    return sectionTotals;
  }

  const opTotals  = addSection('── SAÍDAS OPERACIONAIS', CATS_OPERACIONAL, 'FF2E75B6');
  const empTotals = addSection('── EMPRÉSTIMOS / FINANCIAMENTOS', ['Empréstimos / Financiamentos'], 'FFD6001E');

  const grandTotals = MESES.map((_,i) => opTotals[i] + empTotals[i]);
  const grandRow = wsFC.addRow(['═══ TOTAL SAÍDAS DE CAIXA', ...grandTotals, grandTotals.reduce((a,b)=>a+b,0)]);
  grandRow.font = { bold: true, size: 12 };
  grandRow.eachCell((cell, colNumber) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    if (colNumber > 1) cell.numFmt = numFmt;
  });
  console.log(`✅ FC_Pagos gerado | Total: R$ ${grandTotals.reduce((a,b)=>a+b,0).toLocaleString('pt-BR',{minimumFractionDigits:2})}`);

  await wb.xlsx.writeFile(XLSX_PATH);
  console.log(`\n🎉 Excel salvo com sucesso: ${XLSX_PATH}`);
  console.log(`   Abas adicionadas: Pagos | FC_Pagos`);
}

main().catch(err => { console.error('❌ Erro:', err.message); process.exit(1); });
