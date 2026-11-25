// =============================
// 📊 RELATÓRIO MENSAL
// =============================

function gerarRelatorioMensalWeb(mes, ano) {

  const banco = getAba(CONFIG.ABAS.BANCO);
  const dados = banco.getDataRange().getValues();

  dados.shift();

  const resultado = [];

  dados.forEach(l => {

    const data = new Date(l[3]);

    if (data.getMonth() + 1 != mes || data.getFullYear() != ano) return;

    resultado.push({
      tipo: l[4] || "",
      conta: l[13] || "",
      tr: l[0],
      data: formatarData(l[3]),
      empenho: l[16] || "",
      ugEmpenho: l[17] || "",
      ug: l[7] || "",
      processo: l[5] || "",
      documento: l[8] || "",
      fornecedor: l[9] || "",
      valor: Number(l[14]) || 0,
      beneficiario: l[20] || "",
      cpf: l[21] || "",
      siape: l[22] || ""
    });

  });

  // 🔥 ordenação otimizada
resultado.sort((a, b) => {

  const tipoA = String(a.tipo || "");
  const tipoB = String(b.tipo || "");

  const contaA = String(a.conta || "");
  const contaB = String(b.conta || "");

  if (tipoA !== tipoB) return tipoA.localeCompare(tipoB);

  if (contaA !== contaB) return contaA.localeCompare(contaB);

  const dataA = new Date(a.data.split("/").reverse().join("-"));
  const dataB = new Date(b.data.split("/").reverse().join("-"));

  return dataA - dataB;
});

  return resultado;
}

// =============================
// 📥 EXPORTAR RELATÓRIO EXCEL
// =============================

function exportarRelatorioExcel(mes, ano) {

  const dados = gerarRelatorioMensalWeb(mes, ano);

  if (!dados.length) {
    throw new Error("Nenhum dado encontrado para o período");
  }

  const linhas = [];

  // CABEÇALHO
  linhas.push([
    "TIPO","CONTA","TR","DATA","EMPENHO","UG EMPENHO",
    "UG","PROCESSO","DOCUMENTO","FORNECEDOR",
    "VALOR","BENEFICIARIO","CPF","SIAPE"
  ]);

  dados.forEach(d => {
    linhas.push([
      d.tipo,
      d.conta,
      d.tr,
      d.data,
      d.empenho,
      d.ugEmpenho,
      d.ug,
      d.processo,
      d.documento,
      d.fornecedor,
      d.valor,
      d.beneficiario,
      d.cpf,
      d.siape
    ]);
  });

  // 🔥 cria planilha temporária
  const nova = SpreadsheetApp.create("Relatorio_TEMP");
  const aba = nova.getSheets()[0];

  aba.getRange(1,1,linhas.length,linhas[0].length).setValues(linhas);

  const url = "https://docs.google.com/spreadsheets/d/" +
              nova.getId() + "/export?format=xlsx";

  return url;
}