const CHAVE_REGISTROS = "registros";
const CHAVE_CONFIG = "config";

// Onde o SGD mostra o usuario logado. Fixo: nao precisa de configuracao.
const SGD_PADRAO = "https://sgd.dominiosistemas.com.br/*";
const SGD_SELETOR = "#navbar > p > a > b";

const MOTIVOS = [
  "Cliente desligou antes do atendimento",
  "Ligação muda / sem áudio",
  "Problema técnico no softphone",
  "Estava em pausa / intervalo",
  "Estava em outro atendimento",
  "Queda de conexão ou energia",
  "Chamada duplicada",
  "Trote / engano",
  "Outro (descrever na observação)"
];

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

async function lerConfig() {
  const dados = await chrome.storage.local.get(CHAVE_CONFIG);
  return dados[CHAVE_CONFIG] || {};
}

async function gravarConfig(config) {
  await chrome.storage.local.set({ [CHAVE_CONFIG]: config });
}

function gerarId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
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

// Le telefone / codigo do cliente / horario a partir dos parametros da URL
// aberta pelo screen pop. Aceita varios nomes possiveis de parametro para
// facilitar o encaixe com a URL que o SGD ja usa hoje.
function extrairDadosDaUrl(url) {
  const achado = { telefone: "", codCliente: "", horario: "" };
  if (!url) return achado;

  let params;
  try {
    params = new URL(url).searchParams;
  } catch (e) {
    return achado;
  }

  const pegar = (nomes) => {
    for (const nome of nomes) {
      for (const [chave, valor] of params.entries()) {
        if (chave.toLowerCase() === nome && valor) return valor;
      }
    }
    return "";
  };

  achado.telefone = pegar(["telefone", "phone", "ani", "fone", "numero", "callernumber"]);
  achado.codCliente = pegar(["codcliente", "codigocliente", "codigo", "customerid", "clientid", "idcliente"]);
  achado.horario = pegar(["horario", "hora", "datahora", "timestamp"]);

  return achado;
}

// Limpa o texto capturado na tela do SGD ("Ola, Maria Silva" -> "Maria Silva").
function limparNomeAtendente(texto) {
  if (!texto) return "";
  let limpo = String(texto).replace(/\s+/g, " ").trim();
  limpo = limpo.replace(
    /^(ola|bem-vindo|bem vindo|bem-vinda|usuario|logado como|voce|perfil)[,:\s-]+/i,
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
