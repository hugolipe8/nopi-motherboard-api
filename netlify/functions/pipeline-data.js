const XLSX = require('xlsx');
const fetch = require('node-fetch');

const DROPBOX_URL = 'https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx?rlkey=4px2hpxbg8p6fot2l65bkdamg&dl=1';

const COL = {
  PQPROC:   54,
  AGENCIA:  55,
  DATA_PREV: 56,
  TIPO:     57,
  ID:       58,
  DATA:     59,
  TN:       60,
  REF:      61,
  ENTIDADE: 62,
  FASE:     66,
  VVENDA:   67,
  COMISSAO: 68,
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

function calcComissao(tn, comissao, hasF1, hasF3) {
  if (hasF1) return null; // "-" já recebida
  const tnUp = (tn || '').toUpperCase();
  if (tnUp === 'V1' || tnUp === 'A1') {
    // Pleno de agência — metade para cada consultor
    return hasF3 ? comissao / 4 : comissao / 2;
  } else if (tnUp === 'V2' || tnUp === 'A2') {
    // Partilha imóvel nosso — 1ª linha é nossa
    return hasF3 ? comissao / 4 : comissao / 2;
  } else if (tnUp === 'V3' || tnUp === 'A3') {
    // Partilha imóvel externo — 2ª linha é nossa
    return hasF3 ? comissao / 2 : comissao / 2;
  }
  return comissao / 2;
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

    // Filtrar apenas linhas TIPO=PROC e PQPROC vazio (não fechados)
    const procRows = dataRows.filter(r =>
      toStr(r[COL.TIPO])?.toUpperCase() === 'PROC' &&
      !toStr(r[COL.PQPROC]) // PQPROC vazio = ativo
    );

    // Agrupar por processo (ID)
    const processosMap = {};
    procRows.forEach(r => {
      const id = toStr(r[COL.ID]);
      if (!id) return;
      if (!processosMap[id]) processosMap[id] = [];
      processosMap[id].push(r);
    });

    // Para cada processo, calcular comissão e identificar consultor NOPI
    const processos = [];

    Object.entries(processosMap).forEach(([id, linhas]) => {
      // Linhas de abertura (FASE=A)
      const linhasA = linhas.filter(r => toStr(r[COL.FASE])?.toUpperCase() === 'A');
      if (linhasA.length === 0) return;

      const tn = toStr(linhasA[0][COL.TN])?.toUpperCase() || '';
      const comissaoBase = toNum(linhasA[0][COL.COMISSAO]) || 0;
      const hasF1 = linhas.some(r => toStr(r[COL.FASE])?.toUpperCase() === 'F1');
      const hasF3 = linhas.some(r => toStr(r[COL.FASE])?.toUpperCase() === 'F3');

      // Identificar consultor NOPI
      let consultorNopi = null;
      let agencia = null;

      if (tn === 'V1' || tn === 'A1') {
        // Ambas as linhas são nossas — criar entrada para cada uma
        linhasA.forEach(linha => {
          const comissao = calcComissao(tn, comissaoBase, hasF1, hasF3);
          processos.push({
            consultor: toStr(linha[COL.ENTIDADE]),
            agencia:   toStr(linha[COL.AGENCIA]),
            processo:  id,
            ref:       toStr(linha[COL.REF]),
            data:      toDate(linha[COL.DATA]),
            dataPrev:  toDate(linha[COL.DATA_PREV]),
            preco:     toNum(linha[COL.VVENDA]),
            comissao:  hasF1 ? null : comissao,
            comissaoRecebida: hasF1 ? 'CPCV' : hasF3 ? 'PARCIAL' : null,
            tn,
          });
        });
        return;
      } else if (tn === 'V2' || tn === 'A2') {
        // 1ª linha é nossa
        if (linhasA.length >= 1) {
          consultorNopi = toStr(linhasA[0][COL.ENTIDADE]);
          agencia = toStr(linhasA[0][COL.AGENCIA]);
        }
      } else if (tn === 'V3' || tn === 'A3') {
        // 2ª linha é nossa
        if (linhasA.length >= 2) {
          consultorNopi = toStr(linhasA[1][COL.ENTIDADE]);
          agencia = toStr(linhasA[1][COL.AGENCIA]);
        }
      }

      if (!consultorNopi) return;

      const comissao = calcComissao(tn, comissaoBase, hasF1, hasF3);
      processos.push({
        consultor: consultorNopi,
        agencia,
        processo:  id,
        ref:       toStr(linhasA[0][COL.REF]),
        data:      toDate(linhasA[0][COL.DATA]),
        dataPrev:  toDate(linhasA[0][COL.DATA_PREV]),
        preco:     toNum(linhasA[0][COL.VVENDA]),
        comissao:  hasF1 ? null : comissao,
        comissaoRecebida: hasF1 ? 'CPCV' : hasF3 ? 'PARCIAL' : null,
        tn,
      });
    });

    return {
      statusCode: 200,
      headers: CORS,
      body: JSON.stringify({ processos }),
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: CORS,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
