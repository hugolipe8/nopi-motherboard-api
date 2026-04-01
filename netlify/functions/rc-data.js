/**
 * Netlify Serverless Function — NOPI Motherboard API
 * GET /.netlify/functions/rc-data?mes=4
 */

const fetch = require("node-fetch");
const XLSX  = require("xlsx");

const EXCEL_URL = [
  "https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx",
  "?rlkey=4px2hpxbg8p6fot2l65bkdamg&st=4h2vu72e&dl=1",
].join("");

const MONTH_OFFSETS = [0, 17, 30, 42, 54, 66, 78, 90, 102, 114, 126, 138];
const MONTH_NAMES   = [
  "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
  "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro",
];

const SKIP = new Set(["brg","bcl","bgc","cg","ag","fp","cm"]);
const STOP = new Set(["total geral","cessados"]);

const CORS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

function toInt(v) {
  if (v == null || v === "") return 0;
  const n = parseFloat(String(v));
  return Number.isFinite(n) ? Math.round(n) : 0;
}

function toNum(v) {
  if (v == null || v === "") return 0;
  const n = parseFloat(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

function fmtDate(v) {
  if (v instanceof Date && !isNaN(v)) {
    const p = (n) => String(n).padStart(2, "0");
    return `${p(v.getDate())}/${p(v.getMonth() + 1)}/${v.getFullYear()}`;
  }
  if (typeof v === "number" && v > 0) {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    const p = (n) => String(n).padStart(2, "0");
    return `${p(d.getUTCDate())}/${p(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
  }
  return String(v || "");
}

function json(statusCode, body, extra = {}) {
  return {
    statusCode,
    headers: { ...CORS, "Content-Type": "application/json; charset=utf-8", ...extra },
    body: JSON.stringify(body),
  };
}

// Detecta se a tabela tem estrutura O/R (meses com dados: FAT-O, FAT-R, ANG-O, ANG-R...)
// ou estrutura colapsada (só R: FAT-R, ANG-R, PROP-R...)
// BRG com O/R: off+2=FAT-R, off+4=ANG-R, off+8=CONT-R
// BRG colapsada / BCL / BGC: off+1=FAT-R, off+2=ANG-R, off+4=CONT-R
function hasORStructure(rows, headerRowStart, headerRowEnd, off) {
  for (let r = headerRowStart; r <= headerRowEnd; r++) {
    const row = rows[r] || [];
    for (let c = off + 1; c < off + 6; c++) {
      if (String(row[c] ?? "").trim() === "O") return true;
    }
  }
  return false;
}

function extractAgency(rows, agencyCode, off) {
  const result = [];

  // Encontrar linha da agência dinamicamente a partir da row 50
  let agRowIdx = -1;
  for (let i = 50; i < rows.length; i++) {
    if (String(rows[i][off] ?? "").trim().toUpperCase() === agencyCode) {
      agRowIdx = i;
      break;
    }
  }

  if (agRowIdx === -1) return result;

  // Detectar estrutura O/R nos cabeçalhos acima
  const isOR = hasORStructure(rows, Math.max(0, agRowIdx - 8), agRowIdx - 1, off);

  // Offsets consoante estrutura
  const fatCol  = isOR ? off + 2 : off + 1;
  const angCol  = isOR ? off + 4 : off + 2;
  const contCol = isOR ? off + 8 : off + 4;

  // Ler consultores dinamicamente
  for (let i = agRowIdx + 1; i < rows.length; i++) {
    const row  = rows[i];
    const name = String(row[off] ?? "").trim();
    if (!name) continue;
    const lower = name.toLowerCase();
    if (STOP.has(lower)) break;
    if (SKIP.has(lower)) continue;

    result.push({
      nome:        name,
      angariações: toInt(row[angCol]),
      cpcv:        toInt(row[contCol]),
      faturacao:   toNum(row[fatCol]),
    });
  }

  return result;
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: CORS, body: "" };
  }

  try {
    const res = await fetch(EXCEL_URL, { timeout: 45_000 });
    if (!res.ok) throw new Error(`Dropbox respondeu HTTP ${res.status}`);
    const buf = await res.buffer();

    const wb = XLSX.read(buf, { type: "buffer", cellDates: true });

    const ws = wb.Sheets["RC"];
    if (!ws) throw new Error('Folha "RC" não encontrada no ficheiro Excel.');
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

    const now = new Date();
    const mesParam = event.queryStringParameters?.mes;
    const mi  = mesParam ? parseInt(mesParam) - 1 : now.getMonth();
    const off = MONTH_OFFSETS[mi];

    const BRG = extractAgency(rows, "BRG", off);
    const BCL = extractAgency(rows, "BCL", off);
    const BGC = extractAgency(rows, "BGC", off);

    // ── Folha MOTHER — últimas angariações BRG/ANG/VO ────────────────────────
    const wsMother = wb.Sheets["MOTHER"];
    const motherRows = wsMother
      ? XLSX.utils.sheet_to_json(wsMother, { header: 1, defval: "" })
      : [];
    const ultimasAngariações = motherRows
      .filter(row =>
        String(row[55] ?? "").trim() === "BRG" &&
        String(row[57] ?? "").trim() === "ANG" &&
        String(row[60] ?? "").trim() === "VO"
      )
      .slice(-5)
      .reverse()
      .map(row => ({
        ref:        String(row[61] ?? "").trim(),
        localidade: String(row[58] ?? "").trim(),
        consultor:  String(row[62] ?? "").trim(),
        valor:      toNum(row[67]),
        data:       fmtDate(row[59]),
        tipo:       String(row[65] ?? "").trim(),
      }));

    return json(
      200,
      { mes: mi + 1, mesNome: MONTH_NAMES[mi], ano: now.getFullYear(), BRG, BCL, BGC, ultimasAngariações },
      { "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0" }
    );

  } catch (err) {
    console.error("[rc-data]", err.message);
    return json(500, { erro: err.message });
  }
};
