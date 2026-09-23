// Service worker: concentra o envio para o SharePoint.
//
// Fica aqui (e nao no popup) porque o popup e destruido assim que o
// atendente clica fora dele. Se o envio rodasse no popup, uma gravacao
// em andamento seria abortada no meio e poderia deixar uma aba orfa.

importScripts("common.js", "sharepoint.js");

chrome.runtime.onMessage.addListener((mensagem, remetente, responder) => {
  if (!mensagem || !mensagem.tipo) return;

  if (mensagem.tipo === "sincronizar") {
    sincronizarPendentes()
      .then(responder)
      .catch((e) =>
        responder({ ok: 0, falhas: 0, total: 0, erro: String((e && e.message) || e) })
      );
    return true;
  }

  if (mensagem.tipo === "diagnostico") {
    spTestarConexao()
      .then(responder)
      .catch((e) => responder({ ok: false, mensagem: String((e && e.message) || e) }));
    return true;
  }

  if (mensagem.tipo === "sessao") {
    spSessaoAtual()
      .then((nome) => responder({ nome: nome }))
      .catch(() => responder({ nome: null }));
    return true;
  }

  if (mensagem.tipo === "abrirLogin") {
    spAbrirLogin()
      .then((tabId) => responder({ tabId: tabId }))
      .catch((e) => responder({ erro: String((e && e.message) || e) }));
    return true;
  }
});

// Assim que qualquer pagina do SharePoint terminar de carregar, o atendente
// passou a ter sessao valida. E o momento natural para escoar os pendentes,
// sem ele precisar fazer nada.
chrome.tabs.onUpdated.addListener((tabId, info, aba) => {
  if (info.status !== "complete") return;
  if (!aba || !aba.url || aba.url.indexOf(SP_HOST) !== 0) return;
  sincronizarPendentes().catch(() => { /* tenta de novo na proxima */ });
});
