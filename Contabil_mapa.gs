//**

function gerarDadosMapaRMB(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Lista de contas baseada na sua sequência de 9 dígitos
    const contasLista = [
      "123110101","123110102","123110103","123110104","123110105","123110106","123110107","123110108","123110109","123110199",
      "123110112","123110113","123110114","123110118","123110120","123110121","123110125","123110201","123110301",
      "123110302","123110303","123110402","123110403","123110404","123110405","123110406","123110501","123110502","123110503",
      "123110505","123110506","123110507","123110601","123110701","123110801","123110802","123119901","123119902","123119908","123119909","123119910"
    ];

    const mapaRelatorio = {};
    contasLista.forEach(c => {
      mapaRelatorio[c] = {
        conta: c,
        saldoInicial: 0, 
        entEmpenho: 0,
        entEdital: 0,
        entOutros: 0,
        totalEntradas: 0,
        saiTransferencia: 0,
        saiOutros: 0,
        totalSaidas: 0,
        saldoAtual: 0
      };
    });

    const mesFiltro = parseInt(filtros.mes);
    const anoFiltro = parseInt(filtros.ano);

    // ==========================================
    // 0. BUSCA DO SALDO INICIAL AUTOMÁTICO (DO MÊS ANTERIOR)
    // ==========================================
    const abaMapaRmb = ss.getSheetByName("MAPA_RMB");
    if (abaMapaRmb) {
      // Calcula qual foi o mês anterior para buscar o histórico
      let mesAnterior = mesFiltro - 1;
      let anoAnterior = anoFiltro;
      if (mesAnterior === 0) {
        mesAnterior = 12;
        anoAnterior = anoFiltro - 1;
      }
      const periodoAnteriorTexto = `${String(mesAnterior).padStart(2, '0')}/${anoAnterior}`;
      
      const dadosMapa = abaMapaRmb.getDataRange().getDisplayValues(); // Lê o texto visível ("05/2026")
      const dadosMapaReais = abaMapaRmb.getDataRange().getValues();   // Lê os números puros
      
      for (let i = 1; i < dadosMapa.length; i++) {
        if (dadosMapa[i][0].trim() === periodoAnteriorTexto) {
          const contaHistorico = String(dadosMapa[i][1]).trim();
          if (mapaRelatorio[contaHistorico]) {
            // O saldo ATUAL do mês passado vira o saldo INICIAL do mês corrente (Coluna D / Índice 3)
            mapaRelatorio[contaHistorico].saldoInicial = parseFloat(dadosMapaReais[i][3]) || 0;
          }
        }
      }
    }

    // ==========================================
    // 1. LEITURA E FILTRAGEM DA ABA ENTRADA
    // ==========================================
    const abaEntrada = ss.getSheetByName("ENTRADA");
    if (abaEntrada) {
      const dadosEntrada = abaEntrada.getDataRange().getValues();
      for (let i = 1; i < dadosEntrada.length; i++) {
        const linha = dadosEntrada[i];
        
        // Coluna B (Índice 1) - Data do Registro
        let celulaData = linha[1];
        let dataMov = (celulaData instanceof Date) ? celulaData : new Date(celulaData);
        if (!dataMov || isNaN(dataMov.getTime()) || dataMov.getFullYear() !== anoFiltro || (dataMov.getMonth() + 1) !== mesFiltro) {
          continue;
        }

        // Coluna D (Índice 3) - Tipo de Entrada
        const tipoEntrada = String(linha[3] || "").trim().toUpperCase();
        // Coluna AA (Índice 26) - Conta Contábil vinculada
        const contaSub = String(linha[26] || "").trim();
        // Coluna AD (Índice 29) - Valor total do registro
        let valorTexto = String(linha[29] || "0").replace(/[R$\s]/g, '').trim();
        let valorCalculado = 0;
        
        if (valorTexto.includes(",") && valorTexto.includes(".")) {
          valorCalculado = parseFloat(valorTexto.replace(/\./g, "").replace(",", ".")) || 0;
        } else if (valorTexto.includes(",")) {
          // 🌟 CORRIGIDO: De valorCalculated para valorCalculado
          valorCalculado = parseFloat(valorTexto.replace(",", ".")) || 0; 
        } else {
          valorCalculado = parseFloat(valorTexto) || 0;
        }

        // Se a conta calculada estiver no nosso mapa, realiza o somatório
        if (mapaRelatorio[contaSub]) {
          if (tipoEntrada.indexOf("EMPENHO") !== -1) {
            mapaRelatorio[contaSub].entEmpenho += valorCalculado;
          } else if (tipoEntrada.indexOf("EDITAL INTERNO") !== -1 || tipoEntrada.indexOf("EDITAL") !== -1) {
            mapaRelatorio[contaSub].entEdital += valorCalculado;
          } else {
            mapaRelatorio[contaSub].entOutros += valorCalculado;
          }
        }
      }
    }

    // Consolida somatórios estruturais e calcula o Saldo Atual Real
    const resultadoFinal = Object.values(mapaRelatorio).map(item => {
      item.totalEntradas = item.entEmpenho + item.entEdital + item.entOutros;
      item.totalSaidas = item.saiTransferencia + item.saiOutros;
      // 🌟 FÓRMULA ATUALIZADA: Inicial + Entradas - Saídas
      item.saldoAtual = item.saldoInicial + item.totalEntradas - item.totalSaidas; 
      return item;
    });

    resultadoFinal.sort((a, b) => a.conta.localeCompare(b.conta));
    return { sucesso: true, dados: resultadoFinal };

  } catch (e) {
    return { sucesso: false, message: "Erro ao gerar o Mapa do RMB: " + e.message };
  }
}

/**
 * 💾 NOVA FUNÇÃO: Grava ou Sobrescreve o fechamento do mapa na aba MAPA_RMB
 * Colunas: PERIODO DE REFERENCIA | CONTA | SALDO ATUAL
 */
function salvarDadosMapaRMB(periodo, listaDados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("MAPA_RMB");
    
    // Se a aba não existir, cria e monta o cabeçalho correto com 4 colunas (guardando o inicial)
    if (!sheet) {
      sheet = ss.insertSheet("MAPA_RMB");
      sheet.appendRow(["PERIODO DE REFERENCIA", "CONTA", "SALDO INICIAL", "SALDO ATUAL"]);
      sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#d9ead3");
    }

    // 🌟 ALTERAÇÃO CRÍTICA 1: Pega os valores exatamente como aparecem na tela (Texto puro)
    const rangeExistente = sheet.getDataRange();
    const valoresVisiveis = rangeExistente.getDisplayValues(); // Evita bugs de conversão de data
    const valoresReais = rangeExistente.getValues(); // Mantém os números puros para não quebrar decimais
    
    // Mantém o cabeçalho original
    const novasLinhas = [[valoresVisiveis[0][0], valoresVisiveis[0][1], valoresVisiveis[0][2], valoresVisiveis[0][3]]]; 
    
    // Filtra e mantém os dados de OUTROS meses intactos
    for (let i = 1; i < valoresVisiveis.length; i++) {
      // Se o período visível na célula for DIFERENTE do que estamos salvando, preserva a linha com os números originais
      if (String(valoresVisiveis[i][0]).trim() !== String(periodo).trim()) {
        novasLinhas.push([
          valoresVisiveis[i][0], // Período (Texto)
          valoresVisiveis[i][1], // Conta (Texto)
          valoresReais[i][2],    // Saldo Inicial (Número puro)
          valoresReais[i][3]     // Saldo Atual (Número puro)
        ]);
      }
    }

    // Adiciona os novos registros computados em tela no lote de salvamento
    listaDados.forEach(item => {
      novasLinhas.push([
        String(periodo).trim(),
        String(item.conta).trim(),
        parseFloat(item.saldoInicial) || 0, // Garante que guardamos o ponto de partida
        parseFloat(item.saldoAtual) || 0
      ]);
    });

    // Limpa a planilha e descarrega a matriz atualizada mantendo os outros meses seguros
    sheet.clearContents();
    sheet.getRange(1, 1, novasLinhas.length, 4).setValues(novasLinhas);
    
    // Formata as colunas C (Inicial) e D (Atual) como Moeda Financeira
    if (novasLinhas.length > 1) {
      sheet.getRange(2, 3, novasLinhas.length - 1, 2).setNumberFormat('R$ #,##0.00;R$ (#,##0.00);"-"');
    }

    return { sucesso: true, message: "Fechamento de " + periodo + " gravado com sucesso, preservando o histórico!" };
  } catch (e) {
    return { sucesso: false, message: "Erro ao salvar na planilha: " + e.toString() };
  }
}