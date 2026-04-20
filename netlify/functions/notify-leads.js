/**
 * Netlify Serverless Function — NOPI Lead Notifications
 * Corre de hora em hora via Netlify Scheduled Functions
 * GET /.netlify/functions/notify-leads
 */

const fetch = require("node-fetch");
const { getStore } = require("@netlify/blobs");

// ── Configuração ──────────────────────────────────────────────────────────────
const VEINTIUNO_WORKSPACE_ID = "79c75092-5511-4911-b726-1e7a0079d1bb";
const SUPABASE_URL = "https://brumjtydtlxhooqrrsch.supabase.co";

const ZAPI_INSTANCE = "3F1ED1501D31A3DEC3DFB20DE66F3711";
const ZAPI_TOKEN    = "22BB2C07274EF7734745B11C";
const ZAPI_URL      = `https://api.z-api.io/instances/${ZAPI_INSTANCE}/token/${ZAPI_TOKEN}/send-text`;

// ── MODO DE TESTE ─────────────────────────────────────────────────────────────
// true  = todas as mensagens vão para o teu número (testes)
// false = mensagens vão para o consultor real (produção)
const MODO_TESTE = true;
const NUMERO_TESTE = "351913962809"; // o teu número sem o +

// ── Credenciais 21Online ───────────────────────────────────────────────────────
// Recomendo mover estes valores para variáveis de ambiente no Netlify:
// VEINTIUNO_EMAIL e VEINTIUNO_PASSWORD
const EMAIL    = process.env.VEINTIUNO_EMAIL    || "filipe.moreira@century21.pt";
const PASSWORD = process.env.VEINTIUNO_PASSWORD || "HugoFilipe.2000";

// ── CORS ──────────────────────────────────────────────────────────────────────
const CORS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

function json(statusCode, body) {
  return {
    statusCode,
    headers: { ...CORS, "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify(body),
  };
}

// ── Login 21Online ─────────────────────────────────────────────────────────────
async function login21Online() {
  const res = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, {
    method: "POST",
    headers: { "Content-Type": "application/json", "apikey": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJydW1qdHlkdGx4aG9vcXJyc2NoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NDA5MjMzNzMsImV4cCI6MjA1NjQ5OTM3M30.N0HPIM-5pYa8rNHTRZDpnLblGiS0VLSOFPa_F5LCqeM" },
    body: JSON.stringify({ email: EMAIL, password: PASSWORD }),
  });
  if (!res.ok) throw new Error(`Login falhou: ${res.status}`);
  const data = await res.json();
  return data.access_token;
}

// ── Buscar leads 21Online ──────────────────────────────────────────────────────
async function buscarLeads(token) {
  const res = await fetch(
    `https://21online.app/api/leads?workspaceID=${VEINTIUNO_WORKSPACE_ID}`,
    { headers: { "Authorization": `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error(`Erro ao buscar leads: ${res.status}`);
  return await res.json();
}

// ── Enviar mensagem WhatsApp via Z-API ─────────────────────────────────────────
async function enviarWhatsApp(numeroDestino, mensagem) {
  const numero = MODO_TESTE ? NUMERO_TESTE : numeroDestino.replace(/\D/g, "");
  const res = await fetch(ZAPI_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ phone: numero, message: mensagem }),
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Z-API erro: ${res.status} — ${err}`);
  }
  return await res.json();
}

// ── Formatar mensagem ──────────────────────────────────────────────────────────
function formatarMensagem(lead, modoTeste) {
  const owner     = lead.pipeline_owner || {};
  const contact   = lead.contact || {};
  const mensagem  = lead.message || lead.mensagem || "";
  const imovelRef = lead.reference || lead.imovel_ref || "";
  const fonte     = lead.source || lead.fonte || "";
  const data      = lead.created_at
    ? new Date(lead.created_at).toLocaleString("pt-PT", { timeZone: "Europe/Lisbon" })
    : "";

  const prefixoTeste = modoTeste
    ? `🧪 *[MODO TESTE]* _(mensagem real iria para ${owner.name || "consultor"})_\n\n`
    : "";

  return (
    `${prefixoTeste}` +
    `🏠 *Nova Lead — NOPI*\n\n` +
    `👤 *Cliente:* ${contact.name || "—"}\n` +
    `📧 ${contact.email || "—"}\n` +
    `📞 ${contact.phone || "—"}\n\n` +
    `🏡 *Imóvel:* ${imovelRef || "—"}\n` +
    `🔎 *Fonte:* ${fonte || "—"}\n\n` +
    (mensagem ? `💬 _"${mensagem.slice(0, 200)}${mensagem.length > 200 ? "..." : ""}"_\n\n` : "") +
    `👨‍💼 *Consultor:* ${owner.name || "—"}\n` +
    `📞 ${owner.phone || "—"}\n` +
    `📅 ${data}`
  );
}

// ── Handler principal ──────────────────────────────────────────────────────────
exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: CORS, body: "" };
  }

  try {
    // 1. Ler último timestamp processado do Netlify Blobs
    const store = getStore("nopi-leads");
    let ultimoTimestamp = null;
    try {
      ultimoTimestamp = await store.get("last_lead_timestamp");
    } catch (_) {
      // primeira execução, ainda não existe
    }

    console.log(`[notify-leads] Último timestamp: ${ultimoTimestamp || "nenhum"}`);

    // 2. Login no 21Online
    const accessToken = await login21Online();
    console.log("[notify-leads] Login OK");

    // 3. Buscar leads
    const leads = await buscarLeads(accessToken);
    const todasLeads = Array.isArray(leads) ? leads : (leads.data || leads.leads || []);
    console.log(`[notify-leads] Total de leads: ${todasLeads.length}`);

    // 4. Filtrar leads novas
    const leadsNovas = ultimoTimestamp
      ? todasLeads.filter(l => {
          const ts = l.created_at || l.data || "";
          return ts > ultimoTimestamp;
        })
      : todasLeads.slice(0, 3); // primeira execução: processa as 3 mais recentes como teste

    console.log(`[notify-leads] Leads novas: ${leadsNovas.length}`);

    if (leadsNovas.length === 0) {
      return json(200, { mensagem: "Sem leads novas.", ultimoTimestamp });
    }

    // 5. Ordenar por data crescente e processar
    leadsNovas.sort((a, b) => (a.created_at || "") > (b.created_at || "") ? 1 : -1);

    const resultados = [];
    for (const lead of leadsNovas) {
      const owner  = lead.pipeline_owner || {};
      const numero = owner.phone || "";
      const msg    = formatarMensagem(lead, MODO_TESTE);

      if (!numero && !MODO_TESTE) {
        console.warn(`[notify-leads] Lead sem número de consultor: ${lead.id}`);
        resultados.push({ leadId: lead.id, status: "sem_numero" });
        continue;
      }

      try {
        await enviarWhatsApp(numero, msg);
        console.log(`[notify-leads] Mensagem enviada para ${MODO_TESTE ? NUMERO_TESTE : numero}`);
        resultados.push({ leadId: lead.id, status: "enviado", consultor: owner.name });
      } catch (err) {
        console.error(`[notify-leads] Erro ao enviar: ${err.message}`);
        resultados.push({ leadId: lead.id, status: "erro", erro: err.message });
      }

      // Pequena pausa entre mensagens para não sobrecarregar o Z-API
      await new Promise(r => setTimeout(r, 1000));
    }

    // 6. Guardar novo timestamp (a lead mais recente processada)
    const maisRecente = leadsNovas[leadsNovas.length - 1];
    const novoTimestamp = maisRecente.created_at || maisRecente.data || new Date().toISOString();
    await store.set("last_lead_timestamp", novoTimestamp);
    console.log(`[notify-leads] Novo timestamp guardado: ${novoTimestamp}`);

    return json(200, {
      processadas: leadsNovas.length,
      resultados,
      novoTimestamp,
      modoTeste: MODO_TESTE,
    });

  } catch (err) {
    console.error("[notify-leads] ERRO:", err.message);
    return json(500, { erro: err.message });
  }
};
