const XLSX = require('xlsx');
const fs = require('fs');

const wb = XLSX.readFile('Resumo ORPEC - PLAYERS 2026 (1).xlsx');

function cleanNum(v) {
  if (v == null) return 0;
  if (typeof v === 'number') return v;
  let s = String(v).replace(/[%,]/g, '').trim();
  let n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function findRow(ws, text, col) {
  for (let r = 1; r <= 300; r++) {
    const v = ws[XLSX.utils.encode_cell({ c: col, r })]?.v;
    if (v && String(v).toUpperCase().includes(text.toUpperCase())) return r;
  }
  return 0;
}

function collectSimple(ws, headerRow, colP, colV) {
  let results = [];
  for (let r = headerRow + 2; r < headerRow + 25; r++) {
    let p = ws[XLSX.utils.encode_cell({ c: colP, r })]?.v;
    let v = ws[XLSX.utils.encode_cell({ c: colV, r })]?.v;
    if (!p) break;
    results.push({ player: String(p).trim(), value: cleanNum(v) });
  }
  return results;
}

function findRankCol(ws) {
  for (let c = 8; c <= 12; c++) {
    for (let r = 20; r <= 40; r++) {
      const v = ws[XLSX.utils.encode_cell({ c, r })]?.v;
      if (v && String(v).toUpperCase().includes('TOUCHPOINTS')) return c;
    }
  }
  return 0;
}

function readRankingSection(ws, sectionText, rankCol) {
  let results = [];
  let headRow = findRow(ws, sectionText, rankCol);
  if (!headRow) return results;
  for (let r = headRow + 2; r < headRow + 20; r++) {
    let p = ws[XLSX.utils.encode_cell({ c: rankCol, r })]?.v;
    if (!p || String(p).toUpperCase() === 'PLAYERS') continue;
    results.push({
      player: String(p).trim(),
      seo: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+1, r })]?.v),
      pagespeed: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+2, r })]?.v),
      gmn: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+3, r })]?.v),
      instagram: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+4, r })]?.v),
      facebook: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+5, r })]?.v),
      linkedin: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+6, r })]?.v),
      youtube: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+7, r })]?.v),
      resultado: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+8, r })]?.v)
    });
  }
  return results;
}

// Detect layout: if "Resultado" is at rankCol+8 it's layout9 (Jan), if at rankCol+7 it's layout8 (Fev/Mar)
function detectResultOffset(ws, rankCol) {
  for (let r = 20; r < 80; r++) {
    const v = ws[XLSX.utils.encode_cell({ c: rankCol+8, r })]?.v;
    if (v && String(v).toUpperCase().includes('RESULTADO')) return 8;
    const v2 = ws[XLSX.utils.encode_cell({ c: rankCol+7, r })]?.v;
    if (v2 && String(v2).toUpperCase().includes('RESULTADO')) return 7;
  }
  return 8; // default
}

// Meta sheets
function parseMeta(sheetName) {
  const ws = wb.Sheets[sheetName];
  let result = {};
  ['INSTAGRAM','FACEBOOK','LINKEDIN','YOUTUBE'].forEach(k => {
    let row = findRow(ws, k, 1);
    result[k] = row ? collectSimple(ws, row, 1, 2) : [];
  });
  return result;
}

// SEO_GMN sheets
function parseGMN(sheetName) {
  const ws = wb.Sheets[sheetName];
  let result = { seo: [], pagespeed: [], gmn: [], rankingTouchpoints: [], rankingPerformance: [], resultado: [], percentualDashboard: [] };

  let seo = findRow(ws, 'PERFORMANCE SITE', 1);
  if (seo) result.seo = collectSimple(ws, seo, 1, 2);
  let ps = findRow(ws, 'PAGE SPEED', 1);
  if (ps) result.pagespeed = collectSimple(ws, ps, 1, 2);
  let gmn = findRow(ws, 'PERFORMANCE GMN', 1);
  if (gmn) result.gmn = collectSimple(ws, gmn+1, 1, 2);

  let rankCol = findRankCol(ws);
  if (!rankCol) return result;

  let resOff = detectResultOffset(ws, rankCol);

  // Read touchpoints - values are small (1-10), performance are larger
  result.rankingTouchpoints = readRankingSection(ws, 'TOUCHPOINTS', rankCol);

  // Read performance - find the PERFORMANCE ranking header
  let perfHeadRow = findRow(ws, 'RANKING DOS PLAYERS - PERFORMANCE', rankCol);
  if (perfHeadRow) {
    for (let r = perfHeadRow + 2; r < perfHeadRow + 20; r++) {
      let p = ws[XLSX.utils.encode_cell({ c: rankCol, r })]?.v;
      if (!p || String(p).toUpperCase() === 'PLAYERS') continue;
      let rowObj = {
        player: String(p).trim(),
        seo: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+1, r })]?.v),
        pagespeed: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+2, r })]?.v),
        gmn: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+3, r })]?.v),
        instagram: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+4, r })]?.v),
        facebook: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+5, r })]?.v),
        linkedin: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+6, r })]?.v),
        youtube: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+7, r })]?.v),
        resultado: cleanNum(ws[XLSX.utils.encode_cell({ c: rankCol+8, r })]?.v)
      };
      result.rankingPerformance.push(rowObj);
    }
  }

  // Resultado final - find the simplified result table
  // After performance ranking, find PLAYERS/Resultado header pair
  if (perfHeadRow) {
    for (let r = perfHeadRow + 20; r < perfHeadRow + 35; r++) {
      let p = ws[XLSX.utils.encode_cell({ c: rankCol, r })]?.v;
      if (p && String(p).toUpperCase() !== 'PLAYERS') {
        let v = ws[XLSX.utils.encode_cell({ c: rankCol + 1, r })]?.v;
        if (v != null && cleanNum(v) > 0) {
          result.resultado.push({ player: String(p).trim(), resultado: cleanNum(v) });
        }
      }
    }
  }

  // Percentual Dashboard
  let pdCol = findRankCol(ws) || rankCol;
  for (let c = pdCol - 2; c <= pdCol + 2; c++) {
    for (let r = 50; r <= 75; r++) {
      const v = ws[XLSX.utils.encode_cell({ c, r })]?.v;
      if (v && String(v).toUpperCase().includes('PORCENTAGEM')) {
        let startR = r + 1;
        for (let r2 = startR; r2 < startR + 8; r2++) {
          let platform = ws[XLSX.utils.encode_cell({ c: c, r2 })]?.v;
          if (platform && ['INSTAGRAM','FACEBOOK','LINKEDIN','YOUTUBE','SITE','GMN'].includes(String(platform).toUpperCase().trim())) {
            result.percentualDashboard.push({
              platform: String(platform).toUpperCase().trim(),
              peso: cleanNum(ws[XLSX.utils.encode_cell({ c: c+1, r2 })]?.v),
              notaMax: cleanNum(ws[XLSX.utils.encode_cell({ c: c+2, r2 })]?.v),
              notaPND: cleanNum(ws[XLSX.utils.encode_cell({ c: c+3, r2 })]?.v),
              nota: cleanNum(ws[XLSX.utils.encode_cell({ c: c+4, r2 })]?.v),
              totalPND: cleanNum(ws[XLSX.utils.encode_cell({ c: c+5, r2 })]?.v),
              percentual: cleanNum(ws[XLSX.utils.encode_cell({ c: c+6, r2 })]?.v)
            });
          }
        }
        c = pdCol + 5; // break outer
        break;
      }
    }
  }

  return result;
}

let allData = { meta: {}, seo_gmn: {} };

['META-Jan.26','META-Fev.26','META-Mar.26'].forEach(name => {
  allData.meta[name] = parseMeta(name);
});

['SEO_GMN-Jan.26','SEO_GMN-Fev.26','SEO_GMN-Mar.26'].forEach(name => {
  allData.seo_gmn[name] = parseGMN(name);
});

fs.writeFileSync('dashboard_data.json', JSON.stringify(allData, null, 2));

console.log('=== EXTRACTION SUMMARY ===');
for (let k of Object.keys(allData.meta)) {
  let d = allData.meta[k];
  console.log(`${k}: IG=${d.INSTAGRAM?d.INSTAGRAM.length:0} FB=${d.FACEBOOK?d.FACEBOOK.length:0} LI=${d.LINKEDIN?d.LINKEDIN.length:0} YT=${d.YOUTUBE?d.YOUTUBE.length:0}`);
}
for (let k of Object.keys(allData.seo_gmn)) {
  let d = allData.seo_gmn[k];
  console.log(`${k}: SEO=${d.seo.length} PS=${d.pagespeed.length} GMN=${d.gmn.length} RT=${d.rankingTouchpoints.length} RP=${d.rankingPerformance.length} RES=${d.resultado.length} PD=${d.percentualDashboard.length}`);
}
