const XLSX = require('xlsx');
const fetch = require('node-fetch');

const DROPBOX_URL = 'https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx?rlkey=4px2hpxbg8p6fot2l65bkdamg&dl=1';

// Índices das colunas relevantes (0-indexed)
// Confirmados via análise da folha MOTHER
const COL = {
  AGENCIA:    55,
  DATA_PREV:  56,
  TIPO:       57,
  ID:         58,
  DATA:       59,
  TN:         60,
  REF:        61,
  ENTIDADE:   62,
  TENTIDADE:  65,
  FASE:       66,
  VVENDA:     67,
  COMISSAO:   68,
};

const toNum = (v) => {
  if (v == null || v === '') return null;
  const n = parseFloat(String(v));
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : null;
};

const toStr = (v) => v != null && v !== '' ? String(v).trim() : null;

const toDate = (v) => {
  if (!v) return null;
  if (v instanceof Date) return v.toISOString().split('T')[0];
  if (typeof v === 'number') {
    try {
      const d = XLSX.SSF.parse_date_code(v);
      return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
    } catch { return null; }
  }
  return String(v).split('T')[0];
};

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Content-Type': 'application/json',
};

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers: CORS, body: '' };
  }

  try {
    const response = await fetch(DROPBOX_URL, { timeout: 45000 });
    const buffer = await response.arrayBuffer();

    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const sheet = workbook.Sheets['MOTHER'];

    // Ler por array de arrays (header:1) para aceder por índice de coluna
    // e evitar confusão com nomes de colunas duplicados
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

    // Ignorar linha de cabeçalho (row 0)
    const dataRows = rows.slice(1);

    const consultores = dataRows
      .filter(r => toStr(r[COL.TIPO])?.toUpperCase() === 'REC')
      .filter(r => toStr(r[COL.ENTIDADE]) && toStr(r[COL.ENTIDADE]) !== 'NOPI')
      .map(r => ({
        nome:              toStr(r[COL.ENTIDADE]),
        agencia:           toStr(r[COL.AGENCIA]),
        objetivoFaturacao: toNum(r[COL.COMISSAO]),
        dataEntrada:       toDate(r[COL.DATA_PREV]),
      }));

    const angariações = dataRows
      .filter(r =>
        toStr(r[COL.TN])?.toUpperCase() === 'VO' &&
        toStr(r[COL.FASE])?.toUpperCase() === 'C'
      )
      .map(r => ({
        consultor:  toStr(r[COL.ENTIDADE]),
        agencia:    toStr(r[COL.AGENCIA]),
        referencia: toStr(r[COL.REF]),
        localidade: toStr(r[COL.ID]),
        tipoImovel: toStr(r[COL.TENTIDADE]),
        preco:      toNum(r[COL.VVENDA]),
        comissao:   toNum(r[COL.COMISSAO]),
        data:       toDate(r[COL.DATA]),
      }));

    return {
      statusCode: 200,
      headers: CORS,
      body: JSON.stringify({ consultores, angariações }),
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: CORS,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
