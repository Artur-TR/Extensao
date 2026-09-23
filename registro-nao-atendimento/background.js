// Service worker: captura as chamadas do screen pop e concentra o envio.
//
// O envio fica aqui (e nao no popup) porque o popup e destruido assim que o
// atendente clica fora dele. Se o envio rodasse no popup, uma gravacao
// em andamento seria abortada no meio e poderia deixar uma aba orfa.

importScripts("common.js", "sharepoint.js");

const ALARME_REENVIO = "reenvio";
const REENVIO_MINUTOS = 5;

// ---------- captura do screen pop ----------

// O Genesys abre uma guia NOVA do SGD a cada chamada, antes mesmo de atender.
// So guias novas contam, para a navegacao normal no SGD nao virar "chamada".
const abasNovas = new Map();
let filaCaptura = Promise.resolve(null);

// Serializa as gravacoes na lista de chamadas: varios eventos da mesma guia
// chegam quase juntos e cada um le e regrava a lista inteira.
function enfileirar(fn) {
  const tarefa = filaCaptura.then(fn);
  filaCaptura = tarefa.catch(() => null);
  return tarefa;
}

function capturar(url, tabId) {
  const quando = Date.now();
  return enfileirar(() => registrarChamadaCapturada(url, tabId, quando));
}

function avaliarAbaNova(tabId, url) {
  const criadaEm = abasNovas.get(tabId);
  if (!criadaEm) return;
  if (Date.now() - criadaEm > 60000) {
    abasNovas.delete(tabId);
    return;
  }
  // A primeira URL com dados vale; redirecionamentos depois (login, home) nao.
  if (urlTemDadosDeChamada(url)) {
    abasNovas.delete(tabId);
    capturar(url, tabId);
  }
}

chrome.tabs.onCreated.addListener((aba) => {
  abasNovas.set(aba.id, Date.now());
  avaliarAbaNova(aba.id, aba.pendingUrl || aba.url);
});

chrome.tabs.onRemoved.addListener((tabId) => abasNovas.delete(tabId));

chrome.tabs.onUpdated.addListener((tabId, info, aba) => {
  if (info.url) avaliarAbaNova(tabId, info.url);

  // Assim que qualquer pagina do SharePoint terminar de carregar, o atendente
  // passou a ter sessao valida. E o momento natural para escoar os pendentes.
  if (info.status === "complete" && aba && aba.url && aba.url.indexOf(SP_HOST) === 0 &&
      !spAbasProprias.has(tabId)) {
    sincronizarSeLivre().catch(() => { /* tenta de novo na proxima */ });
  }
});

// ---------- reenvio automatico ----------

// Com pendentes e sessao valida, tenta de novo a cada 5 min. Se continuar
// falhando (ex.: sem permissao), espaca as tentativas ate 80 min, para nao
// ficar abrindo guia do SharePoint em segundo plano o tempo todo.
async function tentarEnvioAutomatico() {
  const temPendente = (await lerRegistros()).some((r) => !r.enviado);
  if (!temPendente) return;

  const dados = await chrome.storage.local.get(CHAVE_AUTO_ENVIO);
  const estado = dados[CHAVE_AUTO_ENVIO] || { falhas: 0, proxima: 0 };
  if (Date.now() < estado.proxima) return;
  if (!(await spSessaoAtual())) return;

  const resultado = await sincronizarPendentes();
  if (resultado.falhas) {
    const falhas = estado.falhas + 1;
    const espera = REENVIO_MINUTOS * 60000 * Math.pow(2, Math.min(falhas, 4));
    await chrome.storage.local.set({ [CHAVE_AUTO_ENVIO]: { falhas, proxima: Date.now() + espera } });
  } else if (estado.falhas) {
    await chrome.storage.local.set({ [CHAVE_AUTO_ENVIO]: { falhas: 0, proxima: 0 } });
  }
}

chrome.alarms.get(ALARME_REENVIO).then((alarme) => {
  if (!alarme) chrome.alarms.create(ALARME_REENVIO, { periodInMinutes: REENVIO_MINUTOS });
});

chrome.alarms.onAlarm.addListener((alarme) => {
  if (alarme.name === ALARME_REENVIO) tentarEnvioAutomatico().catch(() => {});
});

// ---------- contador de pendentes no icone ----------

async function atualizarBadge() {
  const pendentes = (await lerRegistros()).filter((r) => !r.enviado).length;
  await chrome.action.setBadgeBackgroundColor({ color: "#b45309" });
  await chrome.action.setBadgeText({ text: pendentes ? String(pendentes) : "" });
  await chrome.action.setTitle({
    title: pendentes
      ? "Registrar nao atendimento (" + pendentes + " por enviar)"
      : "Registrar nao atendimento"
  });
}

chrome.storage.onChanged.addListener((mudancas, area) => {
  if (area === "local" && mudancas[CHAVE_REGISTROS]) atualizarBadge().catch(() => {});
});

chrome.runtime.onStartup.addListener(() => atualizarBadge().catch(() => {}));
chrome.runtime.onInstalled.addListener(() => atualizarBadge().catch(() => {}));

// ---------- mensagens do popup e do historico ----------

function responderCom(promessa, responder, seErro) {
  promessa.then(responder).catch((e) => responder(seErro(String((e && e.message) || e))));
  return true;
}

chrome.runtime.onMessage.addListener((mensagem, remetente, responder) => {
  if (!mensagem || !mensagem.tipo) return;

  switch (mensagem.tipo) {
    case "sincronizar":
      return responderCom(sincronizarPendentes(), responder,
        (erro) => ({ ok: 0, falhas: 0, total: 0, erro }));

    case "diagnostico":
      return responderCom(spTestarConexao(), responder,
        (erro) => ({ ok: false, mensagem: erro }));

    case "prepararLista":
      return responderCom(spPrepararLista(), responder,
        (erro) => ({ ok: false, mensagem: traduzirErroSp(erro) }));

    case "sessao":
      return responderCom(spSessaoAtual().then((nome) => ({ nome })), responder,
        () => ({ nome: null }));

    case "abrirLogin":
      return responderCom(spAbrirLogin().then((tabId) => ({ tabId })), responder,
        (erro) => ({ erro }));

    // O popup manda a guia ativa: cobre o caso de a guia do screen pop ter
    // aberto antes de a extensao ser instalada ou recarregada.
    case "capturarAba":
      return responderCom(
        capturar(mensagem.url, mensagem.tabId).then((c) => ({ id: c ? c.id : null })),
        responder,
        () => ({ id: null })
      );

    case "marcarChamada":
      return responderCom(
        enfileirar(() => marcarChamadaRegistrada(mensagem.chamadaId, mensagem.registroId))
          .then(() => ({ ok: true })),
        responder,
        () => ({ ok: false })
      );
  }
});
