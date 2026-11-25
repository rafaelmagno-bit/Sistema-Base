// =============================
// 🔄 CONVERSÃO DE TOMBOS (BI-DIRECIONAL) - COLUNAS CORRIGIDAS
// =============================

function converterTomboLote(texto) {
  // Separa o texto por quebras de linha e remove linhas em branco
  const tombos = texto.split("\n")
    .map(t => normalizarTombo(t))
    .filter(t => t);

  if (tombos.length === 0) return [];

  // Puxa a base de dados unificada da aba TOMBOS
  const dados = getDadosAba(CONFIG.ABAS.TOMBOS);

  if (!dados || dados.length <= 1) {
    throw new Error("A aba TOMBOS está vazia ou não pôde ser lida.");
  }

  const mapaSipacParaSiads = {};
  const mapaSiadsParaSipac = {};

  // Alimenta os dicionários de busca nas duas direções com base na ordem informada:
  // A (0) = tombo_SIADS | E (4) = tombo_SIPAC
  dados.slice(1).forEach(l => {
    const siads = normalizarTombo(l[0]); // 🔥 Coluna A (tombo_SIADS)
    const sipac = normalizarTombo(l[4]); // 🔥 Coluna E (tombo_SIPAC)

    if (sipac && siads) {
      mapaSipacParaSiads[sipac] = siads;
      mapaSiadsParaSipac[siads] = sipac;
    } else if (sipac) {
      mapaSipacParaSiads[sipac] = "Não associado";
    } else if (siads) {
      mapaSiadsParaSipac[siads] = "Não associado";
    }
  });

  // Percorre os tombos digitados e monta o relatório inteligente
  return tombos.map(t => {
    let siadsEncontrado = "-";
    let sipacEncontrado = "-";
    let statusResultado = "NÃO ENCONTRADO";

    if (mapaSipacParaSiads[t] !== undefined) {
      // O usuário digitou um tombo SIPAC válido
      sipacEncontrado = t;
      siadsEncontrado = mapaSipacParaSiads[t];
      statusResultado = "OK (VIA SIPAC)";
    } else if (mapaSiadsParaSipac[t] !== undefined) {
      // O usuário digitou um tombo SIADS válido
      siadsEncontrado = t;
      sipacEncontrado = mapaSiadsParaSipac[t];
      statusResultado = "OK (VIA SIADS)";
    }

    return {
      original: t,
      siads: siadsEncontrado,
      sipac: sipacEncontrado,
      status: statusResultado
    };
  });
}

function normalizarTombo(valor) {
  return String(valor || "")
    .trim()
    .replace(/\s+/g, "")   
    .replace(/\.0$/, "");  
}