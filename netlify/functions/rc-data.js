const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));
const XLSX = require('xlsx');

const DROPBOX_URL = 'https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx?rlkey=4px2hpxbg8p6fot2l65bkdamg&dl=1';

const MONTH_OFFSETS = { 1:0, 2:17, 3:30, 4:42, 5:54, 6:66, 7:78, 8:90, 9:102, 10:114, 11:126, 12:138 };
const IGNORE_NAMES = ['cg', 'ag', 'fp', 'cm'];
const STOP_WORDS = ['total geral', 'cessados'];

const AGENCIES = ['BRG', 'BCL', 'BGC'];

function cellVal(sheet, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = sheet[addr];
  if (!cell) return null;
  return cell.v ?? null;
}

function numVal(sheet, r, c) {
  const v = cellVal(sheet, r, c);
  if (v === null || v === undefined || v === '') return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function strVal(sheet, r, c) {
  const v = cellVal(sheet, r, c);
  return v !== null && v !== undefined ? String(v).trim() : '';
}

function findAgencyRow(sheet, off, agencyCode) {
  // Search from row 50 downwards
  for (let r = 50; r < 300; r++) {
    const val = strVal(sheet, r, off);
    if (val.toUpperCase() === agencyCode) return r;
  }
  return null;
}

function readAgency(sheet, off, agencyRow, isBRG) {
  const consultores = [];
  let r = agencyRow + 1;
  while (r < agencyRow + 100) {
    const nome = strVal(sheet, r, off);
    if (!nome) { r++; continue; }
    const nomeLower = nome.toLowerCase();
    if (STOP_WORDS.some(w => nomeLower.includes(w))) break;
    if (IGNORE_NAMES.includes(nomeLower)) { r++; continue; }

    let ang, cpcv;
    if (isBRG) {
      ang  = numVal(sheet, r, off + 4);
      cpcv = numVal(sheet, r, off + 8);
    } else {
      ang  = numVal(sheet, r, off + 2);
      cpcv = numVal(sheet, r, off + 4);
    }

    consultores.push({ nome: nome.toUpperCase(), angariações: ang, cpcv });
    r++;
  }
  return consultores;
}

exports.handler = async (event) => {
  const headers = {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  const now = new Date(new Date().toLocaleString('en-US', { timeZone: 'Europe/Lisbon' }));
  const mesParam = event.queryStringParameters?.mes;
  const mes = mesParam ? parseInt(mesParam, 10) : now.getMonth() + 1;

  const off = MONTH_OFFSETS[mes];
  if (off === undefined) {
    return { statusCode: 400, headers, body: JSON.stringify({ error: 'Mês inválido (1-12)' }) };
  }

  try {
    const resp = await fetch(DROPBOX_URL);
    if (!resp.ok) throw new Error(`Dropbox fetch failed: ${resp.status}`);
    const buffer = await resp.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'buffer' });

    const sheet = workbook.Sheets['RC'];
    if (!sheet) throw new Error('Folha RC não encontrada no ficheiro');

    const result = { mes };

    for (const agency of AGENCIES) {
      const agencyRow = findAgencyRow(sheet, off, agency);
      if (agencyRow === null) {
        result[agency] = [];
        continue;
      }
      result[agency] = readAgency(sheet, off, agencyRow, agency === 'BRG');
    }

    return { statusCode: 200, headers, body: JSON.stringify(result) };
  } catch (err) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: err.message }) };
  }
};
