/**
 * Converte valores monetários ou numéricos para Float.
 * Handles: "1.387,56", "1387,56", "R$ 1.387,56", números puros, etc.
 */
function parseValorMonetario(valorBruto) {
  if (typeof valorBruto === "number") return isNaN(valorBruto) ? 0 : valorBruto;
  if (!valorBruto) return 0;

  let vStr = String(valorBruto).trim();
  if (!vStr) return 0;

  // Remove R$ e espaços
  vStr = vStr.replace(/R\$/g, "").replace(/\s/g, "");

  if (vStr.includes(",") && vStr.includes(".")) {
    // Padrão BR com milhar: 1.387,56
    vStr = vStr.replace(/\./g, "").replace(",", ".");
  } else if (vStr.includes(",") && !vStr.includes(".")) {
    // Padrão BR sem milhar: 1387,56
    vStr = vStr.replace(",", ".");
  }

  const num = parseFloat(vStr);
  return isNaN(num) ? 0 : num;
}

/**
 * Busca todas as entradas de um determinado mês/ano para o relatório RMB.
 * Agrupa os dados por Tipo de Entrada e ordena as linhas internamente por Conta.
 * @param {Object} filtros - Contém { mes: "01" a "12", ano: "2026" }
 */
function getRelatorioRMB(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ENTRADA");
    if (!aba) return { sucesso: false, message: "Aba 'ENTRADA' não encontrada." };

    const range = aba.getDataRange();
    const dados = range.getValues();
    // Captura os dados exatamente como estão formatados/visíveis na planilha
    const dadosTexto = range.getDisplayValues();

    if (dados.length <= 1) return { sucesso: true, dadosAgrupados: {} };

    const mesFiltro = parseInt(filtros.mes) - 1; // JavaScript usa meses de 0 a 11
    const anoFiltro = parseInt(filtros.ano);

    const linhasFiltradas = [];
    
    for (let i = 1; i < dados.length; i++) {
      const r = dados[i];
      const rTexto = dadosTexto[i]; // Linha correspondente em formato de texto visível
      
      let dataReg = r[COL.DATA_REGISTRO];
      
      if (!(dataReg instanceof Date)) {
        continue; // Ignora linhas sem data válida
      }

      if (dataReg.getFullYear() === anoFiltro && dataReg.getMonth() === mesFiltro) {
        const dataFormatada = Utilities.formatDate(dataReg, Session.getScriptTimeZone(), "dd/MM/yyyy");

        // Tratamento numérico blindado para o Total do Item
        let valorBruto = r[COL.TOTAL_ITEM_COM_FRETE];
        let totalItem = 0;

        if (typeof valorBruto === "number") {
          totalItem = valorBruto;
        } else if (valorBruto) {
          let vStr = String(valorBruto).trim();
          
          if (vStr.includes("R$") || (vStr.includes(",") && vStr.includes("."))) {
            vStr = vStr.replace(/R\$/g, "").replace(/\s/g, "");
            vStr = vStr.replace(/\./g, ""); // Remove pontos de milhar
            vStr = vStr.replace(",", ".");  // Troca vírgula decimal por ponto
          } else if (vStr.includes(",") && !vStr.includes(".")) {
            vStr = vStr.replace(",", ".");
          }
          
          totalItem = parseFloat(vStr) || 0;
        }

        // Pega o número do documento/NF a partir do texto visível na planilha
        let docNF = String(rTexto[COL.NF_NUMERO] || "").trim();

        linhasFiltradas.push({
          conta: String(rTexto[COL.CONTA] || "").trim(),
          sequencial: String(rTexto[COL.SEQUENCIAL] || "").trim(),
          dataRegistro: dataFormatada,
          empenho: String(rTexto[COL.EMPENHO] || "").trim(),
          ugEmpenho: String(rTexto[COL.UG_EMPENHO] || "").trim(),
          ug: String(rTexto[COL.UG] || "").trim(),
          processo: String(rTexto[COL.PROCESSO] || "").trim(),
          documentoNF: docNF || "Não Informada",
          fornecedor: String(rTexto[COL.FORNECEDOR] || "").trim(),
          totalItem: totalItem,
          beneficiario: String(rTexto[COL.BENEFICIARIO] || "").trim(),
          cpf: String(rTexto[COL.CPF] || "").trim(),
          siape: String(rTexto[COL.SIAPE] || "").trim(),
          tipoEntrada: String(rTexto[COL.TIPO] || "").trim() || "NÃO ESPECIFICADO"
        });
      }
    }

    // Ordenar todas as linhas filtradas pela Conta
    linhasFiltradas.sort((a, b) => a.conta.localeCompare(b.conta, undefined, { numeric: true, sensitivity: 'base' }));

    // Agrupar os resultados por Tipo de Entrada
    const dadosAgrupados = {};
    linhasFiltradas.forEach(item => {
      const tipo = item.tipoEntrada;
      if (!dadosAgrupados[tipo]) {
        dadosAgrupados[tipo] = [];
      }
      dadosAgrupados[tipo].push(item);
    });

    return { sucesso: true, dadosAgrupados: dadosAgrupados };

  } catch (e) {
    console.log("Erro ao gerar Relatório RMB: " + e.message);
    return { sucesso: false, message: e.message };
  }
}