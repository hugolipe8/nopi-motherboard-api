const XLSX = require('xlsx');
const fetch = require('node-fetch');

const DROPBOX_URL = 'https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx?rlkey=4px2hpxbg8p6fot2l65bkdamg&dl=1';

const toNum = (v) => {
  if (v == null || v === '') return null;
  const n = parseFloat(String(v));
  return Number.isFinite(n) ? n : null;
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

    // Não usar cellDates:true em conjunto com raw:true — conflito
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const sheet = workbook.Sheets['MOTHER'];

    // raw:false para que os números sejam lidos corretamente
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: false });

    const consultores = rows
      .filter(r => toStr(r['TIPO'])?.toUpperCase() === 'REC')
      .filter(r => toStr(r['ENTIDADE']) && toStr(r['ENTIDADE']) !== 'NOPI')
      .map(r => ({
        nome: toStr(r['ENTIDADE']),
        agencia: toStr(r['AGENCIA']),
        objetivoFaturacao: toNum(r['COMISSAO']),
        dataEntrada: toDate(r['DATA PREV']),
      }));

    const angariações = rows
      .filter(r =>
        toStr(r['TN'])?.toUpperCase() === 'VO' &&
        toStr(r['FASE'])?.toUpperCase() === 'C'
      )
      .map(r => ({
        consultor: toStr(r['ENTIDADE']),
        agencia: toStr(r['AGENCIA']),
        referencia: toStr(r['REF']),
        localidade: toStr(r['ID']),
        tipoImovel: toStr(r['TENTIDADE']),
        preco: toNum(r['VVENDA']),
        comissao: toNum(r['COMISSAO']),
        data: toDate(r['DATA']),
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
