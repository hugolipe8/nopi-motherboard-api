const XLSX = require('xlsx');
const fetch = require('node-fetch');

const DROPBOX_URL = 'https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx?rlkey=4px2hpxbg8p6fot2l65bkdamg&dl=1';

const toNum = (v) => typeof v === 'number' ? v : null;
const toStr = (v) => v != null ? String(v).trim() : null;
const toDate = (v) => {
  if (!v) return null;
  if (v instanceof Date) return v.toISOString().split('T')[0];
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
  return String(v);
};

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };

  try {
    const response = await fetch(DROPBOX_URL);
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true, cellFormula: false });

    const sheet = workbook.Sheets['MOTHER'];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });

    const consultores = rows
      .filter(r => toStr(r['TIPO'])?.toUpperCase() === 'REC')
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
      headers,
      body: JSON.stringify({ consultores, angariações }),
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
