// pdf_parser_comum.js — Parser compartilhado de relatórios Gdoor (a pagar / a receber)
// Node.js + pdfreader
// Layout: tabela com múltiplos registros agrupados na mesma linha (y) lado a lado (x)

const {PdfReader} = require('./node_modules/pdfreader');

const DATE_RE  = /^\d{2}\/\d{2}\/\d{4}$/;
const CODE_RE  = /^0{1,2}\d{4,5}$/;
const VALUE_RE = /^[\d.]+,\d{2}$/;

// rowParserFn(texts: string[]) => array de registros [doc, emis, venc, cod, nome, hist, valor, ...extra]
function parsePDF(pdfPath, rowParserFn) {
  const parseFn = rowParserFn || parseRow;
  return new Promise((resolve, reject) => {
    const pages = [];
    let curPage = null;

    new PdfReader().parseFileItems(pdfPath, (err, item) => {
      if (err) { reject(err); return; }
      if (!item) {
        const records = [];
        for (const page of pages) {
          const rowMap = {};
          for (const it of page) {
            const yKey = Math.round(it.y * 10);
            if (!rowMap[yKey]) rowMap[yKey] = [];
            rowMap[yKey].push(it);
          }
          for (const yKey of Object.keys(rowMap).sort((a,b)=>a-b)) {
            const row = rowMap[yKey].sort((a,b) => a.x - b.x);
            const texts = row.map(r => r.text.trim()).filter(Boolean);
            parseFn(texts).forEach(r => records.push(r));
          }
        }
        // remove duplicatas (mesmo doc+cod+valor)
        const seen = new Set();
        const uniq = records.filter(r => {
          const k = `${r[0]}|${r[3]}|${r[6]}`;
          if (seen.has(k)) return false;
          seen.add(k); return true;
        });
        resolve(uniq);
        return;
      }
      if (item.page) { curPage = []; pages.push(curPage); }
      if (item.text && curPage) curPage.push({x: item.x, y: item.y, text: item.text});
    });
  });
}

// Parser padrão (usado por gerar_apagar.js): [doc, emis, venc, cod, nome, hist, valor]
function parseRow(texts) {
  const dates  = texts.filter(t => DATE_RE.test(t));
  const codes  = texts.filter(t => CODE_RE.test(t));
  const values = texts.filter(t => VALUE_RE.test(t));

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
    records.push([
      docCandidates[i] || '',
      emissaoDates[i]  || '',
      vencDates[i],
      parseInt(codes[i]),
      '',   // nome preenchido depois pelo DePara
      '',   // historico não extraído
      parseFloat(String(values[i]).replace(/\./g,'').replace(',','.'))
    ]);
  }
  return records;
}

module.exports = { parsePDF, parseRow, DATE_RE, CODE_RE, VALUE_RE };
