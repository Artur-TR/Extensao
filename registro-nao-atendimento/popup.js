const el = (id) => document.getElementById(id);

let chamadas = [];
let atendenteOrigem = "manual";
let dadosEditados = false;
let enviando = false;

function abrirHistorico() {
  chrome.tabs.create({ url: chrome.runtime.getURL("historico.html") });
}

// O envio roda no service worker para nao ser abortado se o popup fechar.
function pedirSincronizacao() {
  return chrome.runtime.sendMessage({ tipo: "sincronizar" });
}

async function pedirSessao() {
  try {
    const resposta = await chrome.runtime.sendMessage({ tipo: "sessao" });
    return (resposta && resposta.nome) || null;
  } catch (e) {
    return null;
  }
}

// Sem sessao, os registros continuam sendo salvos, mas ficam pendentes.
function mostrarBarraLogin(semSessao) {
  el("barraLogin").style.display = semSessao ? "flex" : "none";
  if (semSessao) {
    el("textoLogin").textContent = "Sem sessão do Microsoft 365. Os registros ficarão pendentes.";
  }
}

async function entrarNaMicrosoft() {
  await chrome.runtime.sendMessage({ tipo: "abrirLogin" });
  window.close();
}

function mostrarAtendente(nome, origem, etiquetaTexto, etiquetaClasse, dica) {
  const campo = el("atendente");
  campo.value = nome || "";
  campo.readOnly = origem === "SGD";
  atendenteOrigem = origem;
  el("origemAtendente").textContent = etiquetaTexto;
  el("origemAtendente").className = "etiqueta " + etiquetaClasse;
  el("dicaAtendente").textContent = dica || "";
}

// Prioridade: SGD aberto agora > ultimo nome lido do SGD > digitado antes >
// nome da conta Microsoft 365.
async function carregarAtendente(sessao) {
  const doSgd = await obterAtendenteDoSgd();
  if (doSgd) {
    mostrarAtendente(doSgd, "SGD", "lido do SGD", "ok");
    return;
  }

  const dica = "SGD não encontrado. Abra o SGD em uma guia para o nome vir automático.";
  const config = await lerConfig();
  if (config.sgdUltimoValor) {
    mostrarAtendente(config.sgdUltimoValor, "SGD (memorizado)", "último do SGD", "ok", dica);
    return;
  }
  if (config.atendenteManual) {
    mostrarAtendente(config.atendenteManual, "manual", "manual", "alerta", dica);
    return;
  }
  const nomeM365 = await sessao;
  if (nomeM365) {
    mostrarAtendente(nomeM365, "Microsoft 365", "Microsoft 365", "ok", dica);
    return;
  }
  mostrarAtendente("", "manual", "manual", "alerta", dica);
}

function tempoDesde(iso) {
  const minutos = Math.round((Date.now() - new Date(iso).getTime()) / 60000);
  if (minutos < 1) return "agora mesmo";
  if (minutos < 60) return "há " + minutos + " min";
  const horas = Math.floor(minutos / 60);
  return "há " + horas + " h" + (minutos % 60 ? " " + (minutos % 60) + " min" : "");
}

function rotuloChamada(c) {
  const partes = [formatarHora(c.horario), c.telefone || "sem telefone"];
  if (c.codCliente) partes.push("cód " + c.codCliente);
  return partes.join(" · ") + (c.registroId ? "  ✓ registrada" : "");
}

// Se o atendente abriu o popup em cima da guia do screen pop, essa e a
// chamada que ele quer registrar.
async function capturarAbaAtiva() {
  try {
    const [aba] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!aba || !aba.url || aba.url.indexOf(SGD_ORIGEM) !== 0) return null;
    const resposta = await chrome.runtime.sendMessage({ tipo: "capturarAba", url: aba.url, tabId: aba.id });
    return (resposta && resposta.id) || null;
  } catch (e) {
    return null;
  }
}

async function carregarChamadas(selecionarId) {
  chamadas = await lerChamadas();
  const select = el("chamada");
  select.innerHTML = "";

  chamadas.forEach((c, i) => {
    const opcao = document.createElement("option");
    opcao.value = c.id;
    opcao.textContent = (i === 0 ? "Última: " : "") + rotuloChamada(c);
    select.appendChild(opcao);
  });

  const manual = document.createElement("option");
  manual.value = "";
  manual.textContent = "Outra chamada (preencher à mão)";
  select.appendChild(manual);

  const existe = chamadas.some((c) => c.id === selecionarId);
  select.value = existe ? selecionarId : (chamadas[0] ? chamadas[0].id : "");
  aplicarChamada();
}

function chamadaSelecionada() {
  return chamadas.find((c) => c.id === el("chamada").value) || null;
}

function aplicarChamada() {
  const c = chamadaSelecionada();
  const dica = el("dicaChamada");
  dadosEditados = false;
  dica.className = "dica";

  if (c) {
    el("telefone").value = c.telefone || "";
    el("codCliente").value = c.codCliente || "";
    el("idInteracao").value = c.idInteracao || "";
    el("dataHora").value = paraValorDatetimeLocal(new Date(c.horario));
    dica.textContent = "Capturada do SGD " + tempoDesde(c.capturadaEm) + ".";
    if (c.registroId) {
      dica.textContent += " Esta chamada já foi registrada.";
      dica.className = "dica alerta";
    }
    return;
  }

  el("telefone").value = "";
  el("codCliente").value = "";
  el("idInteracao").value = "";
  el("dataHora").value = paraValorDatetimeLocal(new Date());
  dica.textContent = chamadas.length
    ? "Preencha os dados da chamada."
    : "Nenhuma chamada capturada do SGD ainda. Preencha os dados à mão.";
}

// Diz ao gestor o quanto confiar nos dados: vieram do screen pop sem
// alteracao, foram corrigidos pelo atendente ou digitados do zero.
function origemDados() {
  if (!chamadaSelecionada()) return "Manual";
  return dadosEditados ? "SGD (editado)" : "SGD (automático)";
}

function montarMotivos() {
  const caixa = el("motivos");
  MOTIVOS.forEach((motivo, i) => {
    const rotulo = document.createElement("label");
    rotulo.className = "motivo";

    const radio = document.createElement("input");
    radio.type = "radio";
    radio.name = "motivo";
    radio.value = motivo;
    radio.addEventListener("change", aoEscolherMotivo);

    const tecla = document.createElement("span");
    tecla.className = "tecla";
    tecla.textContent = String(i + 1);

    const texto = document.createElement("span");
    texto.textContent = motivo;

    rotulo.append(radio, tecla, texto);
    caixa.appendChild(rotulo);
  });
}

function motivoSelecionado() {
  const marcado = document.querySelector("input[name=motivo]:checked");
  return marcado ? marcado.value : "";
}

function aoEscolherMotivo() {
  const exige = motivoExigeObservacao(motivoSelecionado());
  el("obsObrigatoria").hidden = !exige;
  if (exige) el("observacao").focus();
}

function tratarTeclas(e) {
  const alvo = e.target;
  const tag = alvo.tagName;
  const digitando = tag === "TEXTAREA" || tag === "SELECT" || (tag === "INPUT" && alvo.type !== "radio");

  if (!digitando && /^[1-9]$/.test(e.key)) {
    const radio = document.querySelectorAll("input[name=motivo]")[Number(e.key) - 1];
    if (radio) {
      e.preventDefault();
      radio.checked = true;
      radio.focus();
      aoEscolherMotivo();
    }
    return;
  }

  if (e.key !== "Enter" || e.shiftKey || tag === "SELECT" || tag === "BUTTON") return;
  if (tag === "TEXTAREA" && !e.ctrlKey) return;
  e.preventDefault();
  el("formulario").requestSubmit();
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

function validar(registro) {
  if (!registro.motivo) {
    return { mensagem: "Escolha o motivo (teclas 1 a " + MOTIVOS.length + ")." };
  }
  if (motivoExigeObservacao(registro.motivo) && !registro.observacao) {
    return { mensagem: "Descreva o motivo na observação.", campo: "observacao" };
  }
  if (!registro.dataHora) {
    return { mensagem: "Informe a data e hora da chamada.", campo: "dataHora" };
  }
  if (!registro.atendente) {
    return { mensagem: "Informe o nome do atendente.", campo: "atendente" };
  }
  return null;
}

async function enviar(evento) {
  evento.preventDefault();
  if (enviando) return;

  const chamada = chamadaSelecionada();
  const dataHora = new Date(el("dataHora").value);
  const registro = {
    id: gerarId(),
    atendente: el("atendente").value.trim(),
    atendenteOrigem: atendenteOrigem,
    telefone: normalizarTelefone(el("telefone").value),
    codCliente: el("codCliente").value.trim(),
    idInteracao: el("idInteracao").value.trim(),
    dataHora: isNaN(dataHora) ? "" : dataHora.toISOString(),
    motivo: motivoSelecionado(),
    observacao: el("observacao").value.trim(),
    origemDados: origemDados(),
    chamadaId: chamada ? chamada.id : null,
    origem: chamada ? chamada.url : "",
    registradoEm: new Date().toISOString(),
    versao: versaoExtensao(),
    enviado: false
  };

  const problema = validar(registro);
  if (problema) {
    mostrarAviso(problema.mensagem, "erro");
    if (problema.campo) el(problema.campo).focus();
    return;
  }

  enviando = true;
  el("btnRegistrar").disabled = true;

  await adicionarRegistro(registro);
  if (chamada) {
    await chrome.runtime.sendMessage({ tipo: "marcarChamada", chamadaId: chamada.id, registroId: registro.id });
  }
  if (atendenteOrigem !== "SGD") {
    await gravarConfig({ ...(await lerConfig()), atendenteManual: registro.atendente });
  }

  mostrarAviso("Registro salvo. Enviando para a lista...", "ok");
  document.querySelectorAll("input[name=motivo]").forEach((r) => { r.checked = false; });
  el("observacao").value = "";
  el("obsObrigatoria").hidden = true;
  await carregarChamadas(chamada ? chamada.id : "");
  await atualizarContador();

  const resultado = await pedirSincronizacao();
  await atualizarContador();
  enviando = false;
  el("btnRegistrar").disabled = false;

  if (resultado && resultado.falhas === 0) {
    mostrarAviso("Registrado e enviado para a lista.", "ok");
    setTimeout(() => window.close(), 1200);
    return;
  }

  mostrarBarraLogin(!(await pedirSessao()));
  const motivo = traduzirErroSp(resultado && resultado.erro).replace(/\.?\s*$/, ".");
  mostrarAviso("Salvo aqui, mas o envio falhou: " + motivo + " Será reenviado sozinho depois.", "erro");
}

async function iniciar() {
  montarMotivos();

  el("formulario").addEventListener("submit", enviar);
  el("chamada").addEventListener("change", () => {
    aplicarChamada();
    if (!chamadaSelecionada()) el("telefone").focus();
  });
  ["telefone", "codCliente", "idInteracao", "dataHora"].forEach((id) =>
    el(id).addEventListener("input", () => { dadosEditados = true; })
  );
  el("linkHistorico").addEventListener("click", abrirHistorico);
  el("btnLogin").addEventListener("click", entrarNaMicrosoft);
  document.addEventListener("keydown", tratarTeclas);

  const sessao = pedirSessao();
  const idDaAbaAtiva = await capturarAbaAtiva();
  await Promise.all([carregarChamadas(idDaAbaAtiva), carregarAtendente(sessao), atualizarContador()]);

  // Reenvia em segundo plano o que ficou pendente de tentativas anteriores.
  sessao.then((nome) => {
    mostrarBarraLogin(!nome);
    if (nome) pedirSincronizacao().then(atualizarContador);
  });
}

document.addEventListener("DOMContentLoaded", iniciar);
