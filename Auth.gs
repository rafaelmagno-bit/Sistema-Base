// =============================
// 🔐 CONTROLE DE ACESSO
// =============================

function getPerfilUsuario() {

  const email = Session.getEffectiveUser().getEmail();

  const dados = getDadosAba("PERFIS");

  const linha = dados.find(l => l[0] === email);

  return linha ? linha[1] : "VISUALIZADOR";
}