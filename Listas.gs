// =============================
// 🔽 LISTAS DINÂMICAS
// =============================

function getListas() {

  const aba = getAba(CONFIG.ABAS.LISTAS);
  const dados = aba.getDataRange().getValues();

  const tipos = [];
  const contas = [];
  const ugs = [];

  dados.slice(1).forEach(l => {
    if (l[0]) tipos.push(l[0]);
    if (l[1]) contas.push(l[1]);
    if (l[2]) ugs.push(l[2]);
  });

  return { tipos, contas, ugs };
}