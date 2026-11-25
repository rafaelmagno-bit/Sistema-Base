// =============================
// ⚡ CACHE DE PLANILHA
// =============================

function getDadosAba(nomeAba) {

  const cache = CacheService.getScriptCache();
  const chave = "aba_" + nomeAba;

  const cached = cache.get(chave);

  if (cached) {
    return JSON.parse(cached);
  }

  const aba = getAba(nomeAba);
  const dados = aba.getDataRange().getValues();

  // 🔥 NÃO CACHEAR ABAS GRANDES
  if (dados.length < 1000) {
    cache.put(chave, JSON.stringify(dados), 300);
  }

  return dados;
}