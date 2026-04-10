const XLSX = require('xlsx');
const fetch = require('node-fetch');

const DROPBOX_URL = 'https://www.dropbox.com/scl/fi/y4i9m6v4q8snd2m3qljoh/Motherboard-2026.xlsx?rlkey=4px2hpxbg8p6fot2l65bkdamg&dl=1';

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };

  try {
    const response = await fetch(DROPBOX_URL);
    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });

    const sheet = workbook.Sheets['MOTHER'];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

    // Consultores (TIPO=REC)
    const consultores = rows
      .filter(r => r['TIPO'] === 'REC')
      .map(r => ({
        nome: r['ENTIDADE'],
        agencia: r['AGENCIA'],
        objetivoFaturacao: r['COMISSAO'],
        dataEntrada: r['DATA PREV'],
      }));

    // Angariações ativas (TN=VO, FASE=c)
    const angariações = rows
      .filter(r =>
        r['TN'] === 'VO' &&
        String(r['FASE'] ?? '').toLowerCase() === 'c'
      )
      .map(r => ({
        consultor: r['ENTIDADE'],
        agencia: r['AGENCIA'],
        referencia: r['REF'],
        localidade: r['ID'],
        tipoImovel: r['TENTIDADE'],
        preco: r['VVENDA'],
        comissao: r['COMISSAO'],
        data: r['DATA'],
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
