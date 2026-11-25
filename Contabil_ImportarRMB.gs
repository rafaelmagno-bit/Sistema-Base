function processarERegistrarRmb(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName("RMB_SIADS");
  
  // Cria a aba caso ela não exista na planilha
  if (!aba) {
    aba = ss.insertSheet("RMB_SIADS");
    aba.appendRow([
      "TIPO", "REFERENCIA", "UG", "CONTA", "SALDO ANTERIOR", 
      "ENTRADA ORÇAMENTÁRIA", "ENTRADA EXTRA ORÇAMENTÁRIA", "ENTRADA ALTERAÇÃO DE CONTA", 
      "SAIDA", "SAIDA ALTERAÇÃO DE CONTA", "DEPRECIAÇÃO", "SALDO ATUAL"
    ]);
  }

  // Converte a string em formato de meses vinda do HTML para MM/AAAA (Ex: 2026-05 -> 05/2026)
  const partesData = payload.referencia.split("-");
  const referenciaFormatada = partesData[1] + "/" + partesData[0];
  const ug = payload.ug.trim();
  
  // Divide o conteúdo do arquivo em linhas textuais
  const linhas = payload.txtContent.split(/\r?\n/);
  
  let linhasParaSalvar = [];
  let contaAtual = null;
  
  // Dicionário temporário para mapear as linhas de valores de cada bloco
  let dadosBloco = {
    "SALDO ANTERIOR": { lancado: 0, depreciado: 0 },
    "ENTRADA ORCAMENTARIA": { lancado: 0, depreciado: 0 },
    "ENTRADA EXTRA ORCAMENTARIA": { lancado: 0, depreciado: 0 },
    "ENTRADA ALTERAÇÃO DE CONTA": { lancado: 0, depreciado: 0 }, // Unifica variações normais de string
    "SAIDA": { lancado: 0, depreciado: 0 },
    "SAIDA ALTERAÇÃO DE CONTA": { lancado: 0, depreciado: 0 },
    "DEPRECIACAO": { lancado: 0, depreciado: 0 },
    "SALDO ATUAL": { lancado: 0, depreciado: 0 }
  };

  // Função auxiliar para converter o padrão monetário do TXT ("22.464.067,69") em Float numérico puro
  function limparNumero(str) {
    if (!str) return 0;
    let n = str.replace(/\./g, "").replace(",", ".").trim();
    return parseFloat(n) || 0;
  }

  // Percorre todo o relatório linha por linha
  for (let i = 0; i < linhas.length; i++) {
    let linha = linhas[i];

    // 1. Identifica a abertura de um bloco de conta contábil (Ex: | 12.311.01.01/14.212.04.00 ...)
    if (linha.includes("|") && (linha.includes("12.") || linha.includes("14.")) && (linha.includes(" - ") || linha.includes("/")) ) {
      
      // Se já tínhamos uma conta rastreada pendente, descarrega o bloco anterior antes de abrir o novo
      if (contaAtual && contaAtual !== "*** TOTAIS ***") {
        empilharDadosConta(linhasParaSalvar, contaAtual, referenciaFormatada, ug, dadosBloco);
      }
      
      // Captura o código da conta limpando espaços extras
      let partes = linha.split("|")[1].split("-")[0].trim();
      // Pega apenas a primeira parte caso venha combinada por barra "/"
      contaAtual = partes.split("/")[0].trim();
      
      // Reinicializa o dicionário de valores para a nova conta
      Object.keys(dadosBloco).forEach(k => { dadosBloco[k] = { lancado: 0, depreciado: 0 }; });
      continue;
    }

    // Identifica o fim das contas reais e o início do rodapé de Totais Gerais do arquivo
    if (linha.includes("*** TOTAIS ***")) {
      if (contaAtual && contaAtual !== "*** TOTAIS ***") {
        empilharDadosConta(linhasParaSalvar, contaAtual, referenciaFormatada, ug, dadosBloco);
      }
      contaAtual = "*** TOTAIS ***"; // Ignora registros futuros
      continue;
    }

    // Se não estiver dentro de um bloco de conta válida ou for os totais, ignora a linha
    if (!contaAtual || contaAtual === "*** TOTAIS ***") continue;

    // 2. Mapeamento de valores das linhas com base nas strings identificadoras do SIADS
    let chaveRmb = null;
    if (linha.includes("SALDO ANTERIOR")) chaveRmb = "SALDO ANTERIOR";
    else if (linha.includes("ENTRADA ORCAMENTARIA")) chaveRmb = "ENTRADA ORCAMENTARIA";
    else if (linha.includes("ENTRADA EXTRA ORCAMENTARIA")) chaveRmb = "ENTRADA EXTRA ORCAMENTARIA";
    else if (linha.includes("ENTRADA ALTERACAO DE CONTA")) chaveRmb = "ENTRADA ALTERAÇÃO DE CONTA";
    else if (linha.includes("SAIDA ALTERACAO DE CONTA")) chaveRmb = "SAIDA ALTERAÇÃO DE CONTA";
    else if (linha.includes("SAIDA") && !linha.includes("ALTERACAO")) chaveRmb = "SAIDA";
    else if (linha.includes("DEPRECIACAO")) chaveRmb = "DEPRECIACAO";
    else if (linha.includes("SALDO ATUAL")) chaveRmb = "SALDO ATUAL";

    if (chaveRmb) {
      // Divide por "|" para isolar a região das duas colunas de valores à direita
      let colunasValores = linha.split("|");
      if (colunasValores.length >= 4) {
        let txtLancado = colunasValores[2] ? colunasValores[2].trim() : "";
        let txtDepreciado = colunasValores[3] ? colunasValores[3].trim() : "";

        dadosBloco[chaveRmb].lancado = limparNumero(txtLancado);
        dadosBloco[chaveRmb].depreciado = limparNumero(txtDepreciado);
      }
    }
  }

  // Descarrega o último bloco de conta lido do laço
  if (contaAtual && contaAtual !== "*** TOTAIS ***") {
    empilharDadosConta(linhasParaSalvar, contaAtual, referenciaFormatada, ug, dadosBloco);
  }

  // Insere em lote todas as linhas geradas na planilha de uma só vez (Performance otimizada)
  if (linhasParaSalvar.length > 0) {
    aba.getRange(aba.getLastRow() + 1, 1, linhasParaSalvar.length, 12).setValues(linhasParaSalvar);
  }

  return "OK";
}

// Organiza as informações lidas estruturando duas linhas físicas na planilha (VALOR LANCADO e DEPRECIACAO ACUM.)
function empilharDadosConta(targetArray, conta, referencia, ug, dados) {
  // Linha 1: VALOR LANCADO
  targetArray.push([
    "VALOR LANCADO", referencia, ug, conta,
    dados["SALDO ANTERIOR"].lancado,
    dados["ENTRADA ORCAMENTARIA"].lancado,
    dados["ENTRADA EXTRA ORCAMENTARIA"].lancado,
    dados["ENTRADA ALTERAÇÃO DE CONTA"].lancado,
    dados["SAIDA"].lancado,
    dados["SAIDA ALTERAÇÃO DE CONTA"].lancado,
    dados["DEPRECIACAO"].lancado,
    dados["SALDO ATUAL"].lancado
  ]);

  // Linha 2: DEPRECIACAO ACUM.
  targetArray.push([
    "DEPRECIACAO ACUM.", referencia, ug, conta,
    dados["SALDO ANTERIOR"].depreciado,
    dados["ENTRADA ORCAMENTARIA"].depreciado,
    dados["ENTRADA EXTRA ORCAMENTARIA"].depreciado,
    dados["ENTRADA ALTERAÇÃO DE CONTA"].depreciado,
    dados["SAIDA"].depreciado,
    dados["SAIDA ALTERAÇÃO DE CONTA"].depreciado,
    dados["DEPRECIACAO"].depreciado,
    dados["SALDO ATUAL"].depreciado
  ]);
}