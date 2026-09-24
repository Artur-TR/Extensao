const CHAVE_REGISTROS = "registros";
const CHAVE_CONFIG = "config";
const CHAVE_CHAMADAS = "chamadas";
const CHAVE_AUTO_ENVIO = "autoEnvio";

// Quantas chamadas capturadas do SGD ficam disponiveis para escolha no popup.
const MAX_CHAMADAS = 30;

// Onde o SGD mostra o usuario logado. Fixo: nao precisa de configuracao.
const SGD_PADRAO = "https://sgd.dominiosistemas.com.br/*";
const SGD_ORIGEM = "https://sgd.dominiosistemas.com.br/";
const SGD_SELETOR = "#navbar > p > a > b";

// Motivos focados nas falhas do Genesys que geram Rona. Curtos de proposito:
// cabem numa linha no popup e nos graficos do gestor. Sao gravados como
// texto na lista, entao mudar esta relacao nao quebra registros antigos.
const MOTIVOS = [
  "Chamada não carregou",
  "Ligações sobrepostas",
  "Genesys congelou",
  "Botão atender falhou",
  "Sem áudio / muda",
  "Queda de rede / VPN",
  "Headset / softphone",
  "Outro (descrever)"
];

function motivoExigeObservacao(motivo) {
  return /^outro/i.test(motivo || "");
}

async function lerRegistros() {
  const dados = await chrome.storage.local.get(CHAVE_REGISTROS);
  return dados[CHAVE_REGISTROS] || [];
}

async function gravarRegistros(lista) {
  await chrome.storage.local.set({ [CHAVE_REGISTROS]: lista });
}

async function adicionarRegistro(registro) {
  const lista = await lerRegistros();
  lista.unshift(registro);
  await gravarRegistros(lista);
}

// Aplica alteracoes por id sobre a versao MAIS RECENTE do armazenamento.
// Quem demora (ex.: envio ao SharePoint) nao pode regravar uma copia antiga
// da lista inteira, senao apaga registros feitos enquanto esperava.
async function aplicarNosRegistros(alteracoesPorId) {
  const lista = await lerRegistros();
  for (const registro of lista) {
    const alteracao = alteracoesPorId[registro.id];
    if (!alteracao) continue;
    for (const [chave, valor] of Object.entries(alteracao)) {
      if (valor === undefined) delete registro[chave];
      else registro[chave] = valor;
    }
  }
  await gravarRegistros(lista);
}

async function lerConfig() {
  const dados = await chrome.storage.local.get(CHAVE_CONFIG);
  return dados[CHAVE_CONFIG] || {};
}

async function gravarConfig(config) {
  await chrome.storage.local.set({ [CHAVE_CONFIG]: config });
}

async function lerChamadas() {
  const dados = await chrome.storage.local.get(CHAVE_CHAMADAS);
  return dados[CHAVE_CHAMADAS] || [];
}

async function gravarChamadas(lista) {
  await chrome.storage.local.set({ [CHAVE_CHAMADAS]: lista.slice(0, MAX_CHAMADAS) });
}

function gerarId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function versaoExtensao() {
  return chrome.runtime.getManifest().version;
}

function paraValorDatetimeLocal(data) {
  const ajustada = new Date(data.getTime() - data.getTimezoneOffset() * 60000);
  return ajustada.toISOString().slice(0, 16);
}

function formatarDataHora(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d)) return iso;
  return d.toLocaleString("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function formatarHora(iso) {
  const d = new Date(iso);
  if (isNaN(d)) return "";
  return d.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
}

// Um formato unico de telefone, para o gestor conseguir agrupar e cruzar.
function normalizarTelefone(valor) {
  const bruto = String(valor || "").trim();
  let d = bruto.replace(/\D/g, "");
  if (d.length >= 12 && d.startsWith("55")) d = d.slice(2);
  if (d.length === 11) return "(" + d.slice(0, 2) + ") " + d.slice(2, 7) + "-" + d.slice(7);
  if (d.length === 10) return "(" + d.slice(0, 2) + ") " + d.slice(2, 6) + "-" + d.slice(6);
  return d || bruto;
}

function interpretarHorario(valor) {
  if (!valor) return null;
  if (/^\d{10}$/.test(valor)) return new Date(Number(valor) * 1000);
  if (/^\d{13}$/.test(valor)) return new Date(Number(valor));
  const d = new Date(valor);
  return isNaN(d) ? null : d;
}

const RE_UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const RE_TELEFONE = /^(tel:)?[\d\s()+.-]{10,20}$/i;

// Le telefone / codigo do cliente / id da interacao / horario a partir dos
// parametros da URL aberta pelo screen pop (query string ou hash). Aceita
// varios nomes de parametro; se nenhum nome bater, reconhece pelo formato
// (UUID = id de conversa do Genesys, 10 a 13 digitos = telefone).
function extrairDadosDaUrl(url) {
  const achado = { telefone: "", codCliente: "", idInteracao: "", horario: "" };
  if (!url) return achado;

  let u;
  try {
    u = new URL(url);
  } catch (e) {
    return achado;
  }

  const pares = [...u.searchParams.entries()];
  const hash = u.hash.slice(1);
  if (hash.includes("=")) {
    const q = hash.includes("?") ? hash.slice(hash.indexOf("?") + 1) : hash;
    pares.push(...new URLSearchParams(q).entries());
  }

  const pegar = (nomes) => {
    for (const nome of nomes) {
      for (const [chave, valor] of pares) {
        if (chave.toLowerCase() === nome && valor) return valor.trim();
      }
    }
    return "";
  };

  achado.telefone = pegar(["telefone", "phone", "ani", "fone", "numero", "callernumber", "remote", "tel"]);
  achado.codCliente = pegar(["codcliente", "codigocliente", "codigo", "customerid", "clientid", "idcliente", "cliente"]);
  achado.idInteracao = pegar([
    "conversationid", "conversation", "interactionid", "idinteracao", "interacao",
    "callid", "idchamada", "ucid", "gcid"
  ]);
  achado.horario = pegar(["horario", "hora", "datahora", "timestamp"]);

  if (!achado.idInteracao) {
    const uuid = pares.find(([, v]) => RE_UUID.test(v));
    if (uuid) achado.idInteracao = uuid[1];
  }
  if (!achado.telefone) {
    const usados = [achado.codCliente, achado.horario, achado.idInteracao];
    const fone = pares.find(([, v]) => {
      const digitos = v.replace(/\D/g, "");
      return !usados.includes(v.trim()) && RE_TELEFONE.test(v.trim()) &&
        digitos.length >= 10 && digitos.length <= 13;
    });
    if (fone) achado.telefone = fone[1];
  }

  if (achado.telefone) achado.telefone = normalizarTelefone(achado.telefone);
  return achado;
}

function urlTemDadosDeChamada(url) {
  if (!url || url.indexOf(SGD_ORIGEM) !== 0) return false;
  const d = extrairDadosDaUrl(url);
  return Boolean(d.telefone || d.codCliente || d.idInteracao);
}

// Guarda a chamada aberta pelo screen pop. O horario da captura e o momento
// em que a guia abriu, ou seja, quando a chamada chegou (antes de atender).
// Devolve a chamada (nova ou ja existente) ou null se a URL nao tem dados.
async function registrarChamadaCapturada(url, tabId, quando) {
  if (!urlTemDadosDeChamada(url)) return null;

  const lista = await lerChamadas();
  const agora = quando || Date.now();

  // Mesma guia com a mesma URL e recarga, nao chamada nova. A janela de 60s
  // cobre o screen pop disparando duas guias para a mesma chamada.
  const existente = lista.find((c) =>
    c.url === url && (c.tabId === tabId || agora - new Date(c.capturadaEm).getTime() < 60000)
  );
  if (existente) return existente;

  const dados = extrairDadosDaUrl(url);
  const horarioUrl = interpretarHorario(dados.horario);
  const chamada = {
    id: gerarId(),
    url: url,
    tabId: tabId,
    capturadaEm: new Date(agora).toISOString(),
    horario: (horarioUrl || new Date(agora)).toISOString(),
    telefone: dados.telefone,
    codCliente: dados.codCliente,
    idInteracao: dados.idInteracao,
    registroId: null
  };

  lista.unshift(chamada);
  await gravarChamadas(lista);
  return chamada;
}

async function marcarChamadaRegistrada(chamadaId, registroId) {
  const lista = await lerChamadas();
  const chamada = lista.find((c) => c.id === chamadaId);
  if (!chamada) return;
  chamada.registroId = registroId;
  await gravarChamadas(lista);
}

// Limpa o texto capturado na tela do SGD ("Ola, Maria Silva" -> "Maria Silva").
function limparNomeAtendente(texto) {
  if (!texto) return "";
  let limpo = String(texto).replace(/\s+/g, " ").trim();
  limpo = limpo.replace(
    /^(ol[aá]|bem-vindo|bem vindo|bem-vinda|usu[aá]rio|logado como|voc[eê]|perfil)[,:\s-]+/i,
    ""
  );
  limpo = limpo.replace(/[,:;|]+$/, "").trim();
  return limpo.slice(0, 80);
}

// Le o atendente direto da barra de navegacao do SGD (usuario logado).
// Retorna null quando o SGD nao esta aberto em nenhuma guia.
async function obterAtendenteDoSgd() {
  try {
    const abas = await chrome.tabs.query({ url: SGD_PADRAO });

    for (const aba of abas) {
      const resultado = await chrome.scripting.executeScript({
        target: { tabId: aba.id },
        func: (seletor) => {
          const alvo = document.querySelector(seletor);
          if (!alvo) return "";
          return (alvo.value || alvo.textContent || "").trim();
        },
        args: [SGD_SELETOR]
      });

      const bruto = resultado && resultado[0] && resultado[0].result;
      if (bruto) {
        const nome = limparNomeAtendente(bruto);
        if (nome) {
          await gravarConfig({ ...(await lerConfig()), sgdUltimoValor: nome });
          return nome;
        }
      }
    }
  } catch (e) {
    return null;
  }

  return null;
}
