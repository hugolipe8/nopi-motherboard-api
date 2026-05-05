const XLSX = require('xlsx');
const fetch = require('node-fetch');

const DROPBOX_URL = 'https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx?rlkey=4px2hpxbg8p6fot2l65bkdamg&dl=1';

const COL = {
  PQPROC:    54,
  AGENCIA:   55,
  DATA_PREV: 56,
  TIPO:      57,
  ID:        58,
  DATA:      59,
  TN:        60,
  REF:       61,
  ENTIDADE:  62,
  FASE:      66,
  VVENDA:    67,
  COMISSAO:  68,
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

function calcComissao(tn, comissaoBase, hasF1, hasF3, numConsultoresUnicos) {
  if (hasF1) return null;
  const t = (tn || '').toUpperCase();
  if (t === 'V1' || t === 'A1' || t === 'P1') {
    const parteConsultor = comissaoBase / (numConsultoresUnicos || 1);
    return hasF3 ? parteConsultor / 2 : parteConsultor;
  }
  if (t === 'V2' || t === 'A2' || t === 'P2') return hasF3 ? comissaoBase / 4 : comissaoBase / 2;
  if (t === 'V3' || t === 'A3' || t === 'P3') return hasF3 ? comissaoBase / 2 : comissaoBase;
  return comissaoBase / 2;
}

function processarLinhas(linhasA, linhas, tn, id, hasF1, hasF3, extraFields) {
  const comissaoBase = toNum(linhasA[0][COL.COMISSAO]) || 0;
  const ref = toStr(linhasA[0][COL.REF]);
  const base = {
    processo: id,
    ref:      ref || null,
    data:     toDate(linhasA[0][COL.DATA]),
    preco:    toNum(linhasA[0][COL.VVENDA]),
    comissaoRecebida: hasF1 ? 'CPCV' : hasF3 ? 'PARCIAL' : null,
    tn,
    ...extraFields,
  };

  const resultado = [];
  const t = tn.toUpperCase();

  if (t === 'V1' || t === 'A1' || t === 'P1') {
    const consultoresUnicos = [...new Set(
      linhasA.map(l => toStr(l[COL.ENTIDADE])).filter(Boolean)
    )];
    const comissao = calcComissao(tn, comissaoBase, hasF1, hasF3, consultoresUnicos.length);
    consultoresUnicos.forEach(consultor => {
      const linha = linhasA.find(l => toStr(l[COL.ENTIDADE]) === consultor);
      resultado.push({ ...base, consultor, agencia: toStr(linha[COL.AGENCIA]), comissao });
    });
  } else if (t === 'V2' || t === 'A2' || t === 'P2') {
    const consultor = toStr(linhasA[0][COL.ENTIDADE]);
    if (!consultor) return resultado;
    resultado.push({
      ...base,
      consultor,
      agencia: toStr(linhasA[0][COL.AGENCIA]),
      comissao: calcComissao(tn, comissaoBase, hasF1, hasF3, 1),
    });
  } else if (t === 'V3' || t === 'A3' || t === 'P3') {
    if (linhasA.length < 2) return resultado;
    const consultor = toStr(linhasA[1][COL.ENTIDADE]);
    if (!consultor) return resultado;
    resultado.push({
      ...base,
      consultor,
      agencia: toStr(linhasA[1][COL.AGENCIA]),
      comissao: calcComissao(tn, comissaoBase, hasF1, hasF3, 1),
    });
  }

  return resultado;
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers: CORS, body: '' };
  }

  try {
    const response = await fetch(DROPBOX_URL, { timeout: 45000 });
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    const sheet = workbook.Sheets['MOTHER'];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    const dataRows = rows.slice(1);

    const TN_ATIVOS   = new Set(['V1','V2','V3','A1','A2','A3']);
    const TN_PERDIDOS = new Set(['P1','P2','P3']);

    // ── Processos Ativos — TIPO=PROC, PQPROC vazio ───────────────────────────
    const procAtivosRows = dataRows.filter(r =>
      toStr(r[COL.TIPO])?.toUpperCase() === 'PROC' &&
      !toStr(r[COL.PQPROC]) &&
      TN_ATIVOS.has(toStr(r[COL.TN])?.toUpperCase())
    );

    const ativosMap = {};
    procAtivosRows.forEach(r => {
      const id = toStr(r[COL.ID]);
      if (!id) return;
      if (!ativosMap[id]) ativosMap[id] = [];
      ativosMap[id].push(r);
    });

    const processos = [];
    Object.entries(ativosMap).forEach(([id, linhas]) => {
      const linhasA = linhas.filter(r => toStr(r[COL.FASE])?.toUpperCase() === 'A');
      if (linhasA.length === 0) return;
      const tn = toStr(linhasA[0][COL.TN])?.toUpperCase() || '';
      const hasF1 = linhas.some(r => toStr(r[COL.FASE])?.toUpperCase() === 'F1');
      const hasF3 = linhas.some(r => toStr(r[COL.FASE])?.toUpperCase() === 'F3');
      const extraFields = { dataPrev: toDate(linhasA[0][COL.DATA_PREV]) };
      processos.push(...processarLinhas(linhasA, linhas, tn, id, hasF1, hasF3, extraFields));
    });

    // ── Processos Perdidos — TIPO=PROC, TN=P1/P2/P3, PQPROC=F ───────────────
    const procPerdidosRows = dataRows.filter(r =>
      toStr(r[COL.TIPO])?.toUpperCase() === 'PROC' &&
      toStr(r[COL.PQPROC])?.toUpperCase() === 'F' &&
      TN_PERDIDOS.has(toStr(r[COL.TN])?.toUpperCase())
    );

    const perdidosMap = {};
    procPerdidosRows.forEach(r => {
      const id = toStr(r[COL.ID]);
      if (!id) return;
      if (!perdidosMap[id]) perdidosMap[id] = [];
      perdidosMap[id].push(r);
    });

    const processosPerdidos = [];
    Object.entries(perdidosMap).forEach(([id, linhas]) => {
      const linhasA = linhas.filter(r => toStr(r[COL.FASE])?.toUpperCase() === 'A');
      if (linhasA.length === 0) return;
      const tn = toStr(linhasA[0][COL.TN])?.toUpperCase() || '';
      const hasF1 = false;
      const hasF3 = false;
      const extraFields = { dataPrev: toDate(linhasA[0][COL.DATA_PREV]) };
      processosPerdidos.push(...processarLinhas(linhasA, linhas, tn, id, hasF1, hasF3, extraFields));
    });

    return {
      statusCode: 200,
      headers: CORS,
      body: JSON.stringify({ processos, processosPerdidos }),
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: CORS,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
