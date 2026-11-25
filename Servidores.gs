// =============================
// 👤 SERVIDORES
// =============================

function getServidores() {

  const dados = getDadosAba(CONFIG.ABAS.SERVIDORES); // ✅ CORRETO

  return dados.slice(1).map(l => ({
    nome: l[1],
    cpf: l[0],
    siape: l[2],
    email: l[7],
    telefone: l[8]
  }));
}