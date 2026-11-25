// =============================
// 🔍 CONSULTAR TR
// =============================

function consultarTR(tr) {

  const banco = getAba(CONFIG.ABAS.BANCO);

  const dados = banco.getDataRange().getValues();
  dados.shift();

  const trBusca = Number(tr);

  const registros = dados.filter(l => Number(l[0]) == trBusca);

  if (registros.length === 0) {
    return { status: "vazio" };
  }

  let itens = [];
  let total = 0;

  registros.forEach(l => {

    const valor = Number(l[14]) || 0;

    itens.push({
      ordem: l[COL.ORDEM],
      qtd: l[COL.QTD],
      desc: l[COL.DESCRICAO],
      conta: l[COL.CONTA],
      valor: valor,
      tombamento: l[COL.TOMBAMENTO],
      cadmat: l[COL.CADMAT] // 🔥 necessário para o TXT
    });

    total += valor;
  });

  const primeiro = registros[0];

  return {
  status: "ok",
  tr: trBusca,
  tipo: primeiro[COL.TIPO],
  processo: primeiro[COL.PROCESSO],
  ug: primeiro[COL.UG],
  fornecedor: primeiro[COL.FORNECEDOR],
  data: primeiro[COL.DATA], // 🔥 AQUI ESTÁ O PONTO PRINCIPAL
  uorg: primeiro[COL.UORG], // já aproveita pro TXT
  itens: itens,
  total: total
};
}

function getItensPorSequencial(sequencial) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("ENTRADA");
  const dados = aba.getDataRange().getValues();
  
  // Mapeamos os dados filtrando pelo sequencial
  return dados.slice(1)
    .filter(linha => String(linha[0]) === String(sequencial))
    .map(linha => ({
      ordem: linha[22],
      cadmat: linha[23],
      descricao: linha[24],
      marca: linha[25],
      qtd: linha[27]
    }));
}

function salvarTombamentos(sequencial, listaTombos) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ENTRADA");
  const dados = aba.getDataRange().getValues();
  
  // Percorre a planilha e onde o Sequencial + Ordem bater, grava o tombo
  // Supondo que você criou uma nova coluna (coluna 31) chamada TOMBAMENTO
  listaTombos.forEach(t => {
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(sequencial) && String(dados[i][22]) === String(t.ordem)) {
        aba.getRange(i + 1, 31).setValue(t.valor); // Coluna 31 = Tombamento
        break;
      }
    }
  });
  
  return "Tombamentos vinculados com sucesso!";
}