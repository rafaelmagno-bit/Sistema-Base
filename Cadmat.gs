// =============================
// 📦 CADMAT
// =============================

// =============================
// 📦 CADMAT COM CACHE REAL
// =============================

function buscarCadmatAutocomplete(termo) {

  if (!termo || termo.length < 2) return [];

  const termoNormalizado = normalizar(termo);

  const cache = CacheService.getScriptCache();
  const cacheKey = "cadmat_" + termoNormalizado;

  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  const dados = getDadosAba(CONFIG.ABAS.CADMAT);

  const resultado = dados.slice(1)
    .filter(l => {
      const codigo = String(l[0] || "").toLowerCase(); // Coluna do Código
      const descricao = normalizar(l[1]);             // Coluna da Descrição
      
      // Retorna verdadeiro se o termo estiver no código OU na descrição
      return codigo.includes(termoNormalizado) || descricao.includes(termoNormalizado);
    })
    .slice(0, 10)
    .map(l => ({
      codigo: l[0],
      material: l[1],
      conta: l[7] || l[4]
    }));

  // 🔥 salva cache por 10 minutos
  if (resultado.length > 0 && resultado.length <= 20) {
    cache.put(cacheKey, JSON.stringify(resultado), 600);
  }

  return resultado;
}

// 🔥 função reutilizável
function normalizar(txt) {
  return String(txt || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

// =============================
// 🔎 CONSULTA CADMAT (BOTÃO)
// =============================

function consultarCadmat(termo) {

  if (!termo) return [];

  const termoNormalizado = normalizar(termo);

  // Certifique-se que o CONFIG.ABAS.CADMAT retorne o nome exato "TOMBOS" ou "CADMAT" conforme seu ecossistema
  const dados = getDadosAba(CONFIG.ABAS.CADMAT);

  return dados.slice(1)
    .filter(l => normalizar(l[1]).includes(termoNormalizado)) // Filtra pela coluna de descrição do material
    .map(l => ({
      codigo: l[0],
      material: l[1],
      natureza: l[2],
      subitem: l[3],
      conta: l[7] || l[4]
    }));
    // 🔥 .slice(0, 50) REMOVIDO para que a paginação no cliente pegue todos os registros possíveis!
}