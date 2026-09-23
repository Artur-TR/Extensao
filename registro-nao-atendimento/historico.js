const el = (id) => document.getElementById(id);
let todos = [];

function filtrar() {
  const de = el("de").value ? new Date(el("de").value + "T00:00:00") : null;
  const ate = el("ate").value ? new Date(el("ate").value + "T23:59:59") : null;
  const motivo = el("filtroMotivo").value;
  const busca = el("busca").value.trim().toLowerCase();

  return todos.filter((r) => {
    const d = new Date(r.dataHora);
    if (de && d < de) return false;
    if (ate && d > ate) return false;
    if (motivo && r.motivo !== motivo) return false;
    if (busca) {
      const campos = [r.telefone, r.codCliente, r.idInteracao, r.atendente, r.observacao]
        .join(" ")
        .toLowerCase();
      if (!campos.includes(busca)) return false;
    }
    return true;
  });
}

function renderizar() {
  const lista = filtrar();
  const corpo = el("corpo");
  corpo.innerHTML = "";

  el("vazio").style.display = lista.length ? "none" : "block";

  const pendentes = todos.filter((r) => !r.enviado).length;
  el("resumo").textContent =
    lista.length + " de " + todos.length + " registro(s) exibido(s)" +
    (pendentes ? "  -  " + pendentes + " ainda nao enviado(s) para a lista" : "");

  for (const r of lista) {
    const tr = document.createElement("tr");
    const celulas = [
      [formatarDataHora(r.dataHora)],
      [r.atendente || ""],
      [r.telefone || ""],
      [r.codCliente || ""],
      [r.idInteracao || "", "id"],
      [r.motivo || ""],
      [r.observacao || ""]
    ];
    for (const [texto, classe] of celulas) {
      const td = document.createElement("td");
      td.textContent = texto;
      if (classe) td.className = classe;
      tr.appendChild(td);
    }

    const tdEnvio = document.createElement("td");
    const marca = document.createElement("span");
    if (r.enviado) {
      marca.className = "envio ok";
      marca.textContent = "enviado";
      marca.title = "Enviado em " + formatarDataHora(r.enviadoEm);
    } else {
      marca.className = "envio pendente";
      marca.textContent = "pendente";
      marca.title = r.erroEnvio ? traduzirErroSp(r.erroEnvio) : "Ainda nao enviado para a lista.";
    }
    tdEnvio.appendChild(marca);
    tr.appendChild(tdEnvio);

    const tdAcao = document.createElement("td");
    const botao = document.createElement("button");
    botao.className = "excluir";
    botao.textContent = "excluir";
    botao.addEventListener("click", () => excluir(r.id));
    tdAcao.appendChild(botao);
    tr.appendChild(tdAcao);

    corpo.appendChild(tr);
  }
}

function mostrarAviso(texto, tipo) {
  const aviso = el("aviso");
  aviso.textContent = texto;
  aviso.className = "aviso " + tipo;
}

async function enviarPendentes() {
  const botao = el("btnEnviarPendentes");
  botao.disabled = true;
  mostrarAviso("Enviando...", "ok");

  const resultado = await chrome.runtime.sendMessage({ tipo: "sincronizar" });
  todos = await lerRegistros();
  renderizar();
  botao.disabled = false;

  if (resultado.total === 0) {
    mostrarAviso("Nao ha registros pendentes.", "ok");
  } else if (resultado.falhas === 0) {
    mostrarAviso(resultado.ok + " registro(s) enviado(s) para a lista.", "ok");
  } else {
    mostrarAviso(
      resultado.ok + " enviado(s), " + resultado.falhas + " com falha. " +
      traduzirErroSp(resultado.erro),
      "erro"
    );
  }
}

async function testarConexao() {
  const botao = el("btnTestarConexao");
  botao.disabled = true;
  mostrarAviso("Testando...", "ok");
  const resultado = await chrome.runtime.sendMessage({ tipo: "diagnostico" });
  botao.disabled = false;
  mostrarAviso(resultado.mensagem, resultado.ok ? "ok" : "erro");
}

async function prepararLista() {
  if (!confirm(
    "Criar no SharePoint a lista '" + SP_LISTA + "' e as colunas que faltarem?\n" +
    "Nada existente e apagado. Requer permissao para editar a lista."
  )) return;
  const botao = el("btnPrepararLista");
  botao.disabled = true;
  mostrarAviso("Configurando a lista...", "ok");
  const resultado = await chrome.runtime.sendMessage({ tipo: "prepararLista" });
  botao.disabled = false;
  mostrarAviso(resultado.mensagem, resultado.ok ? "ok" : "erro");
}

async function entrarNaMicrosoft() {
  mostrarAviso(
    "Abrindo a pagina da lista. Se o Microsoft 365 pedir login, faca a entrada; " +
    "os registros pendentes sobem sozinhos assim que a pagina carregar.",
    "ok"
  );
  await chrome.runtime.sendMessage({ tipo: "abrirLogin" });
}

async function excluir(id) {
  if (!confirm(
    "Excluir este registro desta maquina?\n" +
    "Se ele ja foi enviado, continua na lista do SharePoint."
  )) return;
  todos = todos.filter((r) => r.id !== id);
  await gravarRegistros(todos);
  renderizar();
}

async function apagarTudo() {
  if (!confirm("Apagar TODOS os registros? Exporte o CSV antes, se ainda nao exportou.")) return;
  todos = [];
  await gravarRegistros(todos);
  renderizar();
}

function campoCsv(valor) {
  const texto = (valor === null || valor === undefined) ? "" : String(valor);
  return '"' + texto.replace(/"/g, '""') + '"';
}

function exportarCsv() {
  const lista = filtrar();
  if (!lista.length) {
    alert("Nao ha registros para exportar com os filtros atuais.");
    return;
  }

  const cabecalho = [
    "Data/hora da chamada",
    "Atendente",
    "Origem do nome",
    "Telefone",
    "Codigo do cliente",
    "ID da interacao (Genesys)",
    "Motivo",
    "Observacao",
    "Origem dos dados",
    "Registrado em",
    "Enviado para a lista",
    "Versao da extensao"
  ];

  const linhas = lista.map((r) =>
    [
      formatarDataHora(r.dataHora),
      r.atendente,
      r.atendenteOrigem || "manual",
      r.telefone,
      r.codCliente,
      r.idInteracao,
      r.motivo,
      r.observacao,
      r.origemDados,
      formatarDataHora(r.registradoEm),
      r.enviado ? formatarDataHora(r.enviadoEm) : "pendente",
      r.versao
    ].map(campoCsv).join(";")
  );

  // BOM + separador ";" para o Excel em portugues abrir corretamente
  const conteudo = "\uFEFF" + [cabecalho.map(campoCsv).join(";"), ...linhas].join("\r\n");
  const blob = new Blob([conteudo], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const hoje = new Date().toISOString().slice(0, 10);
  const link = document.createElement("a");
  link.href = url;
  link.download = "nao-atendimentos-" + hoje + ".csv";
  link.click();
  URL.revokeObjectURL(url);
}

function limparFiltros() {
  el("de").value = "";
  el("ate").value = "";
  el("filtroMotivo").value = "";
  el("busca").value = "";
  renderizar();
}

async function iniciar() {
  todos = await lerRegistros();

  // Inclui motivos de registros antigos que sairam da relacao atual.
  const select = el("filtroMotivo");
  const motivos = [...new Set([...MOTIVOS, ...todos.map((r) => r.motivo).filter(Boolean)])];
  for (const motivo of motivos) {
    const opcao = document.createElement("option");
    opcao.value = motivo;
    opcao.textContent = motivo;
    select.appendChild(opcao);
  }

  renderizar();

  ["de", "ate", "filtroMotivo", "busca"].forEach((id) =>
    el(id).addEventListener("input", renderizar)
  );
  el("btnCsv").addEventListener("click", exportarCsv);
  el("btnApagarTudo").addEventListener("click", apagarTudo);
  el("btnLimparFiltros").addEventListener("click", limparFiltros);
  el("btnEnviarPendentes").addEventListener("click", enviarPendentes);
  el("btnTestarConexao").addEventListener("click", testarConexao);
  el("btnPrepararLista").addEventListener("click", prepararLista);
  el("btnEntrarMicrosoft").addEventListener("click", entrarNaMicrosoft);

  // Se algum pendente subiu por conta propria enquanto a pagina estava aberta,
  // a tabela se atualiza sozinha.
  chrome.storage.onChanged.addListener(async (mudancas, area) => {
    if (area !== "local" || !mudancas[CHAVE_REGISTROS]) return;
    todos = await lerRegistros();
    renderizar();
  });
}

document.addEventListener("DOMContentLoaded", iniciar);
