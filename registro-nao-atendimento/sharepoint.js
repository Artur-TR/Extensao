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
// Depois use "Configurar lista" na pagina de historico para criar as colunas.

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

// Colunas da lista. "nome" e o nome interno (o que a API grava); "titulo" e
// o que o gestor ve. O autor do registro nao precisa de coluna: a coluna
// nativa "Criado por" ja traz a conta Microsoft 365 de quem enviou.
const SP_COLUNAS = [
  { nome: "DataHoraChamada", titulo: "Data/hora da chamada", xml: "Type='DateTime' Format='DateTime'", visivel: true },
  { nome: "Motivo", titulo: "Motivo", choice: true, visivel: true },
  { nome: "AtendenteSGD", titulo: "Atendente", xml: "Type='Text'", visivel: true },
  { nome: "CodCliente", titulo: "Cód. cliente", xml: "Type='Text'", visivel: true },
  { nome: "IdInteracao", titulo: "ID da interação (Genesys)", xml: "Type='Text'", visivel: true },
  { nome: "Observacao", titulo: "Observação", xml: "Type='Note' RichText='FALSE' NumLines='4'", visivel: true },
  { nome: "Origem", titulo: "Origem dos dados", xml: "Type='Text'", visivel: true },
  { nome: "RegistradoEm", titulo: "Registrado em", xml: "Type='DateTime' Format='DateTime'", visivel: true },
  { nome: "VersaoExtensao", titulo: "Versão da extensão", xml: "Type='Text'", visivel: false },
  // Indexada: sem indice, o filtro anti-duplicata quebra acima de 5.000 itens.
  { nome: "IdLocal", titulo: "IdLocal", xml: "Type='Text' Indexed='TRUE'", visivel: false }
];

// Uma sincronizacao por vez. Quem chega durante uma em andamento espera ela
// terminar e roda a sua, para o registro recem-salvo entrar no envio.
let spSincronizacaoAtual = null;

// Guias do SharePoint abertas pela propria extensao. O carregamento delas nao
// pode disparar nova sincronizacao, senao uma falha vira laco infinito.
const spAbasProprias = new Set();

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
  if (msg.includes("SEM_ABA") || msg.includes("Cannot access")) {
    return "Nao foi possivel abrir uma pagina do SharePoint. Confirme que voce esta logado no Microsoft 365 neste navegador.";
  }
  if (msg.includes("error page")) {
    return "A pagina do SharePoint nao carregou. Verifique a rede ou se o Zscaler esta conectado.";
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
  spAbasProprias.add(aba.id);
  try {
    await esperarCarregar(aba.id);
  } catch (e) {
    await spFecharAbaPropria(aba.id);
    throw e;
  }
  return { tabId: aba.id, criada: true };
}

async function spFecharAbaPropria(tabId) {
  try { await chrome.tabs.remove(tabId); } catch (e) { /* ja fechada */ }
  spAbasProprias.delete(tabId);
}

async function spExecutarNaAba(func, args) {
  const alvo = await spObterAba();
  try {
    const saida = await chrome.scripting.executeScript({
      target: { tabId: alvo.tabId },
      func: func,
      args: args
    });
    return (saida && saida[0] && saida[0].result) ||
      { erro: "sem resposta da pagina do SharePoint" };
  } finally {
    if (alvo.criada) await spFecharAbaPropria(alvo.tabId);
  }
}

function spCamposDoRegistro(r) {
  return {
    Title: r.telefone || "(sem telefone)",
    DataHoraChamada: r.dataHora,
    Motivo: r.motivo,
    AtendenteSGD: r.atendente || "",
    CodCliente: r.codCliente || "",
    IdInteracao: r.idInteracao || "",
    Observacao: r.observacao || "",
    Origem: r.origemDados || "",
    RegistradoEm: r.registradoEm,
    VersaoExtensao: r.versao || "",
    IdLocal: r.id
  };
}

// Executado DENTRO da pagina do SharePoint. Precisa ser autocontido:
// nao enxerga nada do escopo da extensao.
function spInjecaoGravar(site, nomeLista, registros) {
  const urlLista =
    site + "/_api/web/lists/getbytitle('" + encodeURIComponent(nomeLista) + "')";
  const leitura = { Accept: "application/json;odata=nometadata" };

  return (async () => {
    try {
      const respDigest = await fetch(site + "/_api/contextinfo", {
        method: "POST",
        credentials: "include",
        headers: leitura
      });
      if (!respDigest.ok) return { erro: "contextinfo HTTP " + respDigest.status };
      const digest = (await respDigest.json()).FormDigestValue;

      const respTipo = await fetch(urlLista + "?$select=ListItemEntityTypeFullName", {
        credentials: "include",
        headers: leitura
      });
      if (!respTipo.ok) return { erro: "lista HTTP " + respTipo.status };
      const tipo = (await respTipo.json()).ListItemEntityTypeFullName;

      // Grava so as colunas que existem na lista. Uma coluna faltando nao
      // derruba o envio inteiro; o diagnostico aponta o que falta.
      let campos = null;
      const respCampos = await fetch(urlLista + "/fields?$select=InternalName", {
        credentials: "include",
        headers: leitura
      });
      if (respCampos.ok) {
        campos = (await respCampos.json()).value.map((f) => f.InternalName);
      }

      const resultados = [];
      for (const reg of registros) {
        try {
          // Em reenvios, confere antes se o registro ja subiu, para nao
          // duplicar caso a tentativa anterior tenha gravado sem confirmar.
          if (reg.verificarDuplicata) {
            const respExiste = await fetch(
              urlLista + "/items?$select=Id&$top=1&$filter=IdLocal eq '" +
                encodeURIComponent(reg.id) + "'",
              { credentials: "include", headers: leitura }
            );
            if (respExiste.ok) {
              const dados = await respExiste.json();
              if (dados.value && dados.value.length) {
                resultados.push({ id: reg.id, ok: true, jaExistia: true });
                continue;
              }
            }
          }

          const corpo = { __metadata: { type: tipo } };
          for (const [chave, valor] of Object.entries(reg.campos)) {
            if (!campos || campos.includes(chave)) corpo[chave] = valor;
          }

          const resp = await fetch(urlLista + "/items", {
            method: "POST",
            credentials: "include",
            headers: {
              Accept: "application/json;odata=verbose",
              "Content-Type": "application/json;odata=verbose",
              "X-RequestDigest": digest
            },
            body: JSON.stringify(corpo)
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

      return { digestOk: true, tipo: tipo, campos: campos, resultados: resultados };
    } catch (e) {
      return { erro: String((e && e.message) || e) };
    }
  })();
}

function spGravarLote(registros) {
  return spExecutarNaAba(spInjecaoGravar, [SP_SITE, SP_LISTA, registros]);
}

// Envia tudo que ainda nao subiu. O armazenamento local continua sendo a
// fonte de verdade: a lista e uma copia, e nada se perde se o envio falhar.
async function sincronizarPendentes() {
  while (spSincronizacaoAtual) {
    try { await spSincronizacaoAtual; } catch (e) { /* a proxima tenta de novo */ }
  }
  spSincronizacaoAtual = executarSincronizacao();
  try {
    return await spSincronizacaoAtual;
  } finally {
    spSincronizacaoAtual = null;
  }
}

// Para gatilhos automaticos: se ja ha envio em andamento, nao enfileira outro.
function sincronizarSeLivre() {
  if (spSincronizacaoAtual) return Promise.resolve(null);
  return sincronizarPendentes();
}

async function executarSincronizacao() {
  const pendentes = (await lerRegistros()).filter((r) => !r.enviado);
  if (!pendentes.length) return { ok: 0, falhas: 0, total: 0, erro: null };

  // Registros que ja falharam antes sao conferidos na lista antes de
  // gravar de novo, para nao virarem linha duplicada no relatorio.
  const lote = pendentes.map((r) => ({
    id: r.id,
    campos: spCamposDoRegistro(r),
    verificarDuplicata: (r.tentativas || 0) > 0
  }));

  let resposta;
  try {
    resposta = await spGravarLote(lote);
  } catch (e) {
    resposta = { erro: String((e && e.message) || e) };
  }

  const porId = {};
  for (const r of resposta.resultados || []) porId[r.id] = r;

  const agora = new Date().toISOString();
  const alteracoes = {};
  let ok = 0;
  let falhas = 0;
  let erro = resposta.erro || null;

  for (const registro of pendentes) {
    const res = porId[registro.id];
    const tentativas = (registro.tentativas || 0) + 1;
    if (res && res.ok) {
      alteracoes[registro.id] = { enviado: true, enviadoEm: agora, erroEnvio: undefined, tentativas };
      ok++;
    } else {
      const motivo = (res && res.erro) || resposta.erro || "sem resposta para este registro";
      alteracoes[registro.id] = { erroEnvio: motivo, tentativas };
      erro = motivo;
      falhas++;
    }
  }

  await aplicarNosRegistros(alteracoes);
  return { ok: ok, falhas: falhas, total: pendentes.length, erro: erro };
}

// Executado DENTRO da pagina do SharePoint: cria a lista (se nao existir) e
// as colunas que faltam, com nome interno fixo e titulo amigavel.
function spInjecaoPrepararLista(site, nomeLista, colunas) {
  const urlLista =
    site + "/_api/web/lists/getbytitle('" + encodeURIComponent(nomeLista) + "')";
  const leitura = { Accept: "application/json;odata=nometadata" };

  return (async () => {
    try {
      const respDigest = await fetch(site + "/_api/contextinfo", {
        method: "POST",
        credentials: "include",
        headers: leitura
      });
      if (!respDigest.ok) return { erro: "contextinfo HTTP " + respDigest.status };
      const digest = (await respDigest.json()).FormDigestValue;

      const escrever = (url, corpo, mesclar) => fetch(url, {
        method: "POST",
        credentials: "include",
        headers: Object.assign(
          {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": digest
          },
          mesclar ? { "X-HTTP-Method": "MERGE", "IF-MATCH": "*" } : {}
        ),
        body: JSON.stringify(corpo)
      });

      const log = [];
      const respLista = await fetch(urlLista + "?$select=Id", { credentials: "include", headers: leitura });
      if (respLista.status === 404) {
        const r = await escrever(site + "/_api/web/lists", {
          __metadata: { type: "SP.List" },
          BaseTemplate: 100,
          Title: nomeLista,
          Description: "Registros de nao atendimento (Rona) enviados pela extensao."
        });
        if (!r.ok) return { erro: "criar lista HTTP " + r.status + " - " + (await r.text()).slice(0, 150) };
        log.push("Lista criada.");
      } else if (!respLista.ok) {
        return { erro: "lista HTTP " + respLista.status };
      }

      const respCampos = await fetch(urlLista + "/fields?$select=InternalName", { credentials: "include", headers: leitura });
      if (!respCampos.ok) return { erro: "colunas HTTP " + respCampos.status };
      const existentes = (await respCampos.json()).value.map((f) => f.InternalName);

      for (const c of colunas) {
        if (existentes.includes(c.nome)) {
          log.push(c.titulo + ": ja existia");
          continue;
        }
        // Options 8 = usar o Name como nome interno; +16 = mostrar na visao padrao.
        const r = await escrever(urlLista + "/fields/CreateFieldAsXml", {
          parameters: {
            __metadata: { type: "SP.XmlSchemaFieldCreationInformation" },
            SchemaXml: c.schemaXml,
            Options: c.visivel ? 24 : 8
          }
        });
        if (!r.ok) {
          log.push(c.titulo + ": FALHOU - HTTP " + r.status);
          continue;
        }
        if (c.titulo !== c.nome) {
          await escrever(
            urlLista + "/fields/getbyinternalnameortitle('" + c.nome + "')",
            { __metadata: { type: "SP.Field" }, Title: c.titulo },
            true
          );
        }
        log.push(c.titulo + ": criada");
      }

      await escrever(
        urlLista + "/fields/getbyinternalnameortitle('Title')",
        { __metadata: { type: "SP.Field" }, Title: "Telefone", Required: false },
        true
      );
      log.push("Coluna Title renomeada para Telefone.");

      return { ok: true, log: log };
    } catch (e) {
      return { erro: String((e && e.message) || e) };
    }
  })();
}

function escaparXml(texto) {
  return String(texto)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/'/g, "&apos;")
    .replace(/"/g, "&quot;");
}

function spSchemaXml(coluna) {
  if (coluna.choice) {
    const opcoes = MOTIVOS.map((m) => "<CHOICE>" + escaparXml(m) + "</CHOICE>").join("");
    return "<Field Type='Choice' Format='Dropdown' FillInChoice='TRUE' Name='" + coluna.nome +
      "' DisplayName='" + coluna.nome + "'><CHOICES>" + opcoes + "</CHOICES></Field>";
  }
  return "<Field " + coluna.xml + " Name='" + coluna.nome + "' DisplayName='" + coluna.nome + "' />";
}

async function spPrepararLista() {
  const colunas = SP_COLUNAS.map((c) => ({
    nome: c.nome,
    titulo: c.titulo,
    visivel: c.visivel,
    schemaXml: spSchemaXml(c)
  }));
  const resultado = await spExecutarNaAba(spInjecaoPrepararLista, [SP_SITE, SP_LISTA, colunas]);
  if (resultado.erro) return { ok: false, mensagem: traduzirErroSp(resultado.erro) };
  return { ok: true, mensagem: resultado.log.join("\n") };
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

  // Lote vazio: obtem o digest, o tipo e as colunas da lista sem gravar nada.
  let escritaOk = false;
  let detalheEscrita = "";
  let faltando = null;
  try {
    const teste = await spGravarLote([]);
    if (teste.erro) detalheEscrita = teste.erro;
    else escritaOk = Boolean(teste.digestOk);
    if (teste.campos) {
      faltando = SP_COLUNAS.filter((c) => !teste.campos.includes(c.nome)).map((c) => c.nome);
    }
  } catch (e) {
    detalheEscrita = String((e && e.message) || e);
  }

  linhas.push(
    "4. Escrita pela aba do SharePoint: " +
      (escritaOk ? "OK" : "FALHOU - " + (detalheEscrita || "motivo desconhecido"))
  );
  if (faltando) {
    linhas.push(
      "5. Colunas da lista: " +
        (faltando.length ? "faltando " + faltando.join(", ") + " (use Configurar lista)" : "OK")
    );
  }

  let conclusao;
  if (leiturasOk === 3 && escritaOk) {
    conclusao = faltando && faltando.length
      ? "O envio funciona, mas as colunas faltando nao serao gravadas."
      : "Tudo certo. O envio para a lista deve funcionar.";
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
