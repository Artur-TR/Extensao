const el = (id) => document.getElementById(id);

function abrirHistorico() {
  chrome.tabs.create({ url: chrome.runtime.getURL("historico.html") });
}

// O envio roda no service worker para nao ser abortado se o popup fechar.
function pedirSincronizacao() {
  return chrome.runtime.sendMessage({ tipo: "sincronizar" });
}

// Mostra a barra de login so quando nao ha sessao do Microsoft 365.
// Sem sessao, os registros continuam sendo salvos, mas ficam pendentes.
async function verificarSessao() {
  const barra = el("barraLogin");
  const texto = el("textoLogin");

  let resposta;
  try {
    resposta = await chrome.runtime.sendMessage({ tipo: "sessao" });
  } catch (e) {
    resposta = null;
  }

  if (resposta && resposta.nome) {
    barra.style.display = "none";
    return true;
  }

  texto.textContent = "Sem sessao do Microsoft 365. Os registros ficarao pendentes.";
  barra.style.display = "flex";
  return false;
}

async function entrarNaMicrosoft() {
  await chrome.runtime.sendMessage({ tipo: "abrirLogin" });
  window.close();
}

// Preenche o atendente lendo o usuario logado na barra do SGD. So cai para
// digitacao manual quando o SGD nao esta aberto em nenhuma guia.
async function carregarAtendente() {
  const campo = el("atendente");
  const etiqueta = el("origemAtendente");
  const dica = el("dicaAtendente");

  const doSgd = await obterAtendenteDoSgd();

  if (doSgd) {
    campo.value = doSgd;
    campo.readOnly = true;
    etiqueta.textContent = "lido do SGD";
    etiqueta.className = "etiqueta ok";
    dica.textContent = "";
    window.__atendenteOrigem = "SGD";
    return;
  }

  const config = await lerConfig();
  window.__atendenteOrigem = "manual";
  campo.readOnly = false;
  campo.value = config.atendenteManual || config.sgdUltimoValor || "";
  etiqueta.textContent = "manual";
  etiqueta.className = "etiqueta alerta";
  dica.textContent = "SGD nao encontrado. Abra o SGD em uma guia para o nome vir automatico.";
}

async function atualizarContador() {
  const registros = await lerRegistros();
  const pendentes = registros.filter((r) => !r.enviado).length;
  const base = registros.length === 1 ? "1 registro" : registros.length + " registros";
  el("contador").textContent = pendentes ? base + " (" + pendentes + " por enviar)" : base;
}

function mostrarAviso(texto, tipo) {
  const aviso = el("aviso");
  aviso.textContent = texto;
  aviso.className = "aviso " + tipo;
}

async function iniciar() {
  const select = el("motivo");
  for (const motivo of MOTIVOS) {
    const opcao = document.createElement("option");
    opcao.value = motivo;
    opcao.textContent = motivo;
    select.appendChild(opcao);
  }

  await carregarAtendente();

  el("dataHora").value = paraValorDatetimeLocal(new Date());

  await atualizarContador();

  // Tenta preencher telefone / codigo a partir da guia aberta pelo screen pop
  try {
    const [aba] = await chrome.tabs.query({ active: true, currentWindow: true });
    const dados = extrairDadosDaUrl(aba && aba.url);
    let achouAlgo = false;

    if (dados.telefone) { el("telefone").value = dados.telefone; achouAlgo = true; }
    if (dados.codCliente) { el("codCliente").value = dados.codCliente; achouAlgo = true; }
    if (dados.horario) {
      const d = new Date(dados.horario);
      if (!isNaN(d)) { el("dataHora").value = paraValorDatetimeLocal(d); achouAlgo = true; }
    }
    if (achouAlgo) el("detectado").style.display = "block";

    window.__origem = (aba && aba.url) || "";
  } catch (e) {
    window.__origem = "";
  }

  el("formulario").addEventListener("submit", enviar);
  el("btnHistorico").addEventListener("click", abrirHistorico);
  el("linkHistorico").addEventListener("click", abrirHistorico);
  el("btnLogin").addEventListener("click", entrarNaMicrosoft);

  // Reenvia em segundo plano o que ficou pendente de tentativas anteriores.
  verificarSessao().then((temSessao) => {
    if (temSessao) pedirSincronizacao().then(atualizarContador);
  });
}

async function enviar(evento) {
  evento.preventDefault();

  const registro = {
    id: gerarId(),
    atendente: el("atendente").value.trim(),
    atendenteOrigem: window.__atendenteOrigem || "manual",
    telefone: el("telefone").value.trim(),
    codCliente: el("codCliente").value.trim(),
    dataHora: new Date(el("dataHora").value).toISOString(),
    motivo: el("motivo").value,
    observacao: el("observacao").value.trim(),
    origem: window.__origem || "",
    registradoEm: new Date().toISOString(),
    enviado: false
  };

  if (!registro.motivo) {
    mostrarAviso("Selecione o motivo do nao atendimento.", "erro");
    return;
  }

  await adicionarRegistro(registro);
  if (window.__atendenteOrigem !== "SGD") {
    await gravarConfig({ ...(await lerConfig()), atendenteManual: registro.atendente });
  }

  mostrarAviso("Registro salvo. Enviando para a lista...", "ok");

  el("telefone").value = "";
  el("codCliente").value = "";
  el("observacao").value = "";
  el("motivo").value = "";
  el("dataHora").value = paraValorDatetimeLocal(new Date());

  const registros = await lerRegistros();
  el("contador").textContent =
    registros.length === 1 ? "1 registro" : registros.length + " registros";

  const resultado = await pedirSincronizacao();
  await atualizarContador();
  if (resultado && resultado.falhas === 0) {
    mostrarAviso("Registro salvo e enviado para a lista.", "ok");
  } else {
    await verificarSessao();
    mostrarAviso(
      "Salvo aqui, mas o envio falhou: " + traduzirErroSp(resultado && resultado.erro) +
      " Sera reenviado sozinho depois.",
      "erro"
    );
  }
}

document.addEventListener("DOMContentLoaded", iniciar);
