/**
 * Rastreio de Sacas · Stage Zone — fonte de dados
 * Lê a aba "sacas" (tabela por PACOTE) e devolve as SACAS já com a rota real.
 *
 * Regra (RTG_ROUTE_NAME):
 *   - 2 underlines (ex AM1_87_2)  => é uma SACA; o número da saca está em DE_PARA (600–700).
 *   - 1 underline  (ex AM1_87)    => é a ROTA; a rota real está em DE_PARA (ex B7_AM1, VF3_AM1).
 *   A saca e a rota se ligam pelo nome base (AM1_87) — podem compartilhar o ID CONTAINER.
 *
 * COMO PUBLICAR:
 *  1. Abra a planilha → Extensões → Apps Script → cole este arquivo.
 *  2. Implantar → Nova implantação → App da Web (Executar como: Eu · Acesso: Qualquer pessoa).
 *  3. Copie a URL /exec e cole no app (aba Planilha).
 */

const SHEET_ID = '1hRVnx-G2Kw_ZoHYzCGyju914PrB2fGIKZgSHIjm8GxE';
const TAB_NAME = 'sacas';

function doGet() {
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_NAME);
    if (!sh) return json_({ error: 'aba "' + TAB_NAME + '" não encontrada' });

    const values = sh.getDataRange().getValues();

    // localizar cabeçalho
    let hRow = -1, col = {};
    for (let i = 0; i < Math.min(values.length, 30); i++) {
      const row = values[i].map(c => String(c).trim().toUpperCase());
      if (row.indexOf('ID CONTAINER') !== -1 && row.indexOf('DE_PARA') !== -1) {
        hRow = i; row.forEach((name, idx) => { if (name) col[name] = idx; }); break;
      }
    }
    if (hRow < 0) return json_({ error: 'cabeçalho (ID CONTAINER / DE_PARA) não encontrado' });

    const cC = col['ID CONTAINER'], cD = col['DE_PARA'], cR = col['RTG_ROUTE_NAME'];
    const cY = (col['CYCLE'] !== undefined) ? col['CYCLE'] : -1;
    if (cR === undefined) return json_({ error: 'coluna RTG_ROUTE_NAME não encontrada' });

    const invalida = v => !v || v.indexOf('#') !== -1 || /[.\$\/\[\]]/.test(v) || /^#?N\/?A$/i.test(v);

    // passo 1: mapa nome-base -> rota real (linhas de ROTA: 1 underline)
    const rotaMap = {};
    for (let i = hRow + 1; i < values.length; i++) {
      const rtg = String(values[i][cR]).trim();
      if (!rtg) continue;
      const parts = rtg.split('_');
      if (parts.length === 2) {                 // ROTA
        const real = String(values[i][cD]).trim();
        if (real && !invalida(real)) rotaMap[rtg] = real;
      }
    }

    // passo 2: SACAS (2+ underlines) deduplicadas por número da saca (DE_PARA)
    const seen = {};
    for (let i = hRow + 1; i < values.length; i++) {
      const rtg = String(values[i][cR]).trim();
      if (!rtg) continue;
      const parts = rtg.split('_');
      if (parts.length < 3) continue;           // não é saca
      const sacaNum = String(values[i][cD]).trim();      // número da saca (600–700)
      const cid = String(values[i][cC]).trim();
      if (invalida(sacaNum) || invalida(cid)) continue;
      if (seen[sacaNum]) continue;
      const base = parts.slice(0, -1).join('_');         // AM1_87_2 -> AM1_87
      seen[sacaNum] = {
        saca_number: sacaNum,
        container_id: cid,
        route_name: base,
        real_route: rotaMap[base] || base,               // fallback: nome base
        cycle: cY >= 0 ? String(values[i][cY]).trim() : ''
      };
    }

    return json_({
      updated: new Date().toISOString(),
      count: Object.keys(seen).length,
      sacas: Object.keys(seen).map(k => seen[k])
    });
  } catch (err) {
    return json_({ error: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
