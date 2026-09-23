// Envio dos registros para a lista do Microsoft Lists / SharePoint.
//
// Por que a gravacao e feita de dentro de uma aba do SharePoint:
// o SharePoint aceita LEITURA vinda de qualquer origem, mas recusa ESCRITA
// (HTTP 403) quando a requisicao parte de uma extensao. E uma protecao
// anti-CSRF dele. A saida legitima e executar a gravacao dentro de uma
// pagina do proprio SharePoint, onde a requisicao passa a ser de mesma origem.
//
// A extensao reaproveita uma aba do SharePoint se ja houver alguma aberta.
// Se nao houver, abre uma em segundo plano (sem roubar o foco), grava e fecha.
//
// Para apontar para outra lista, troque SP_SITE e SP_LISTA abaixo
// (e atualize o host_permissions no manifest.json para o novo dominio).

const SP_SITE =
  "https://trten-my.sharepoint.com/personal/caue_costadecarvalho_thomsonreuters_com";
const SP_LISTA = "Registros Ronas";
const SP_HOST = new URL(SP_SITE).origin;

// Pagina leve do proprio SharePoint, usada so para ter uma origem valida.
const SP_URL_APOIO = SP_SITE + "/_api/web/currentuser";

// Pagina da lista. Serve tanto para o atendente ver os registros quanto
// para forcar o login do Microsoft 365 e criar a sessao em cookies.
const SP_URL_PAGINA_LISTA =
  SP_SITE + "/Lists/" + encodeURIComponent(SP_LISTA) + "/AllItems.aspx";

// Evita dois envios simultaneos (popup + historico + sincronizacao automatica).
let spSincronizando = false;

function spUrlLista(sufixo) {
  return (
    SP_SITE +
    "/_api/web/lists/getbytitle('" +
    encodeURIComponent(SP_LISTA) +
    "')" +
    (sufixo || "")
  );
}

function traduzirErroSp(erro) {
  const msg = String((erro && erro.message) || erro || "");
  if (msg.includes("SEM_ABA")) {
    return "Nao foi possivel abrir uma pagina do SharePoint. Confirme que voce esta logado no Microsoft 365 neste navegador.";
  }
  if (msg.includes("Failed to fetch") || msg.includes("NetworkError")) {
    return "Nao foi possivel alcancar o SharePoint. Verifique a rede ou se o Zscaler esta conectado.";
  }
  if (msg.includes("403")) {
    return "O SharePoint recusou a gravacao (403). Sua conta pode nao ter permissao de edicao nesta lista.";
  }
  if (msg.includes("404")) {
    return "A lista nao foi encontrada. Confira o nome em SP_LISTA dentro de sharepoint.js.";
  }
  return msg;
}

// Quem esta logado no Microsoft 365 neste navegador, segundo o SharePoint.
// Devolve null quando nao ha sessao valida em cookies.
async function spSessaoAtual() {
  try {
    const resposta = await fetch(SP_HOST + "/_api/web/currentuser", {
      credentials: "include",
      headers: { Accept: "application/json;odata=nometadata" }
    });
    const ehJson = (resposta.headers.get("content-type") || "").includes("json");
    if (!resposta.ok || !ehJson) return null;
    const dados = await resposta.json();
    return dados.Title || dados.Email || dados.LoginName || null;
  } catch (e) {
    return null;
  }
}

// Abre a pagina da lista. Se nao houver sessao, o Microsoft 365 apresenta
// o login e, ao concluir, os cookies passam a existir para a extensao.
async function spAbrirLogin() {
  const aba = await chrome.tabs.create({ url: SP_URL_PAGINA_LISTA, active: true });
  return aba.id;
}

function esperarCarregar(tabId, limiteMs) {
  const limite = limiteMs || 20000;
  const inicio = Date.now();

  return new Promise((resolve, reject) => {
    const checar = () => {
      chrome.tabs.get(tabId).then(
        (aba) => {
          if (aba.status === "complete") return resolve();
          if (Date.now() - inicio > limite) return reject(new Error("SEM_ABA: tempo esgotado"));
          setTimeout(checar, 300);
        },
        () => reject(new Error("SEM_ABA: a aba foi fechada"))
      );
    };
    checar();
  });
}

// Devolve uma aba do SharePoint utilizavel. Reaproveita uma existente
// sempre que possivel, para nao ficar abrindo e fechando abas.
async function spObterAba() {
  const abas = await chrome.tabs.query({ url: SP_HOST + "/*" });
  const pronta = abas.find((a) => a.status === "complete");
  if (pronta) return { tabId: pronta.id, criada: false };

  const aba = await chrome.tabs.create({ url: SP_URL_APOIO, active: false });
  try {
    await esperarCarregar(aba.id);
  } catch (e) {
    try { await chrome.tabs.remove(aba.id); } catch (e2) { /* ja fechada */ }
    throw e;
  }
  return { tabId: aba.id, criada: true };
}

// Executado DENTRO da pagina do SharePoint. Precisa ser autocontido:
// nao enxerga nada do escopo da extensao.
function spInjecaoGravar(site, nomeLista, registros) {
  const urlLista =
    site + "/_api/web/lists/getbytitle('" + encodeURIComponent(nomeLista) + "')";

  return (async () => {
    try {
      const respDigest = await fetch(site + "/_api/contextinfo", {
        method: "POST",
        credentials: "include",
        headers: { Accept: "application/json;odata=nometadata" }
      });
      if (!respDigest.ok) return { erro: "contextinfo HTTP " + respDigest.status };
      const digest = (await respDigest.json()).FormDigestValue;

      const respTipo = await fetch(urlLista + "?$select=ListItemEntityTypeFullName", {
        credentials: "include",
        headers: { Accept: "application/json;odata=nometadata" }
      });
      if (!respTipo.ok) return { erro: "lista HTTP " + respTipo.status };
      const tipo = (await respTipo.json()).ListItemEntityTypeFullName;

      const resultados = [];
      for (const reg of registros) {
        try {
          // Em reenvios, confere antes se o registro ja subiu, para nao
          // duplicar caso a tentativa anterior tenha gravado sem confirmar.
          if (reg.verificarDuplicata) {
            const respExiste = await fetch(
              urlLista + "/items?$select=Id&$top=1&$filter=IdLocal eq '" +
                encodeURIComponent(reg.id) + "'",
              { credentials: "include", headers: { Accept: "application/json;odata=nometadata" } }
            );
            if (respExiste.ok) {
              const dados = await respExiste.json();
              if (dados.value && dados.value.length) {
                resultados.push({ id: reg.id, ok: true, jaExistia: true });
                continue;
              }
            }
          }

          const resp = await fetch(urlLista + "/items", {
            method: "POST",
            credentials: "include",
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "X-RequestDigest": digest
            },
            body: JSON.stringify({
              __metadata: { type: tipo },
              Title: reg.telefone || "(sem telefone)",
              DataHoraChamada: reg.dataHora,
              CodCliente: reg.codCliente || "",
              Motivo: reg.motivo,
              Observacao: reg.observacao || "",
              AtendenteSGD: reg.atendente || "",
              IdLocal: reg.id
            })
          });

          if (resp.ok) {
            resultados.push({ id: reg.id, ok: true });
          } else {
            const texto = await resp.text();
            resultados.push({
              id: reg.id,
              ok: false,
              erro: "HTTP " + resp.status + " - " + texto.slice(0, 150)
            });
          }
        } catch (e) {
          resultados.push({ id: reg.id, ok: false, erro: String((e && e.message) || e) });
        }
      }

      return { digestOk: true, tipo: tipo, resultados: resultados };
    } catch (e) {
      return { erro: String((e && e.message) || e) };
    }
  })();
}

async function spGravarLote(registros) {
  const alvo = await spObterAba();
  try {
    const saida = await chrome.scripting.executeScript({
      target: { tabId: alvo.tabId },
      func: spInjecaoGravar,
      args: [SP_SITE, SP_LISTA, registros]
    });
    return (saida && saida[0] && saida[0].result) ||
      { erro: "sem resposta da pagina do SharePoint" };
  } finally {
    if (alvo.criada) {
      try { await chrome.tabs.remove(alvo.tabId); } catch (e) { /* ja fechada */ }
    }
  }
}

// Envia tudo que ainda nao subiu. O armazenamento local continua sendo a
// fonte de verdade: a lista e uma copia, e nada se perde se o envio falhar.
async function sincronizarPendentes() {
  if (spSincronizando) return { ok: 0, falhas: 0, total: 0, erro: null, ocupado: true };
  spSincronizando = true;

  try {
    const lista = await lerRegistros();
    const pendentes = lista.filter((r) => !r.enviado);
    if (!pendentes.length) return { ok: 0, falhas: 0, total: 0, erro: null };

    // Registros que ja falharam antes sao conferidos na lista antes de
    // gravar de novo, para nao virarem linha duplicada no relatorio.
    const lote = pendentes.map((r) => ({
      id: r.id,
      telefone: r.telefone,
      dataHora: r.dataHora,
      codCliente: r.codCliente,
      motivo: r.motivo,
      observacao: r.observacao,
      atendente: r.atendente,
      verificarDuplicata: (r.tentativas || 0) > 0
    }));

    for (const registro of pendentes) {
      registro.tentativas = (registro.tentativas || 0) + 1;
    }

    let resposta;
    try {
      resposta = await spGravarLote(lote);
    } catch (e) {
      await gravarRegistros(lista);
      return {
        ok: 0,
        falhas: pendentes.length,
        total: pendentes.length,
        erro: String((e && e.message) || e)
      };
    }

    if (resposta.erro) {
      await gravarRegistros(lista);
      return { ok: 0, falhas: pendentes.length, total: pendentes.length, erro: resposta.erro };
    }

    const porId = {};
    for (const r of resposta.resultados || []) porId[r.id] = r;

    let ok = 0;
    let falhas = 0;
    let erro = null;

    for (const registro of pendentes) {
      const res = porId[registro.id];
      if (res && res.ok) {
        registro.enviado = true;
        registro.enviadoEm = new Date().toISOString();
        delete registro.erroEnvio;
        ok++;
      } else {
        registro.erroEnvio = (res && res.erro) || "sem resposta para este registro";
        erro = registro.erroEnvio;
        falhas++;
      }
    }

    await gravarRegistros(lista);
    return { ok: ok, falhas: falhas, total: pendentes.length, erro: erro };
  } finally {
    spSincronizando = false;
  }
}

// Diagnostico passo a passo: leitura direta e escrita pela aba do SharePoint.
async function spDiagnostico() {
  const leituras = [
    { nome: "1. Sessao no Microsoft 365", url: SP_HOST + "/_api/web/currentuser" },
    { nome: "2. Acesso ao site da lista", url: SP_SITE + "/_api/web/currentuser" },
    { nome: "3. Leitura da lista", url: spUrlLista("?$select=ListItemEntityTypeFullName") }
  ];

  const linhas = [];
  let leiturasOk = 0;

  for (const passo of leituras) {
    try {
      const resposta = await fetch(passo.url, {
        credentials: "include",
        headers: { Accept: "application/json;odata=nometadata" }
      });
      const ehJson = (resposta.headers.get("content-type") || "").includes("json");
      if (resposta.ok && ehJson) {
        leiturasOk++;
        linhas.push(passo.nome + ": OK");
      } else {
        linhas.push(passo.nome + ": FALHOU - HTTP " + resposta.status);
      }
    } catch (e) {
      linhas.push(passo.nome + ": FALHOU - " + String((e && e.message) || e));
    }
  }

  // Lote vazio: obtem o digest e o tipo da lista sem gravar nada.
  let escritaOk = false;
  let detalheEscrita = "";
  try {
    const teste = await spGravarLote([]);
    if (teste.erro) detalheEscrita = teste.erro;
    else escritaOk = Boolean(teste.digestOk);
  } catch (e) {
    detalheEscrita = String((e && e.message) || e);
  }

  linhas.push(
    "4. Escrita pela aba do SharePoint: " +
      (escritaOk ? "OK" : "FALHOU - " + (detalheEscrita || "motivo desconhecido"))
  );

  let conclusao;
  if (leiturasOk === 3 && escritaOk) {
    conclusao = "Tudo certo. O envio para a lista deve funcionar.";
  } else if (leiturasOk === 3 && !escritaOk) {
    conclusao =
      "A leitura funciona, mas a gravacao pela aba do SharePoint falhou. " +
      "Verifique se voce tem permissao de edicao na lista.";
  } else if (leiturasOk === 0) {
    conclusao =
      "Nenhuma chamada passou. Provavelmente nao ha sessao do Microsoft 365 neste navegador, " +
      "ou a rede/Zscaler esta bloqueando.";
  } else {
    conclusao =
      "Parte das leituras falhou. Confira SP_SITE e SP_LISTA dentro de sharepoint.js.";
  }

  return { linhas: linhas, conclusao: conclusao, ok: leiturasOk === 3 && escritaOk };
}

async function spTestarConexao() {
  const resultado = await spDiagnostico();
  return {
    ok: resultado.ok,
    mensagem: resultado.linhas.join("\n") + "\n\n" + resultado.conclusao
  };
}
