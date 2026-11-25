/**
 * Gera os dados analíticos e consolidados para o Relatório de Movimentações RMB.
 */
function gerarRelatorioMovimentacoesRMB(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaMov = ss.getSheetByName("DOCUMENTO_MOVIMENTO");
    if (!abaMov) return { sucesso: false, message: "Aba 'DOCUMENTO_MOVIMENTO' não encontrada." };

    const dados = abaMov.getDataRange().getValues();
    if (dados.length <= 1) return { sucesso: true, analitico: [], resumo: [] };

    const mesFiltro = parseInt(filtros.mes);
    const anoFiltro = parseInt(filtros.ano);

    const analitico = [];
    const mapaResumo = {}; 

    for (let i = 1; i < dados.length; i++) {
      const r = dados[i];

      // Coluna C (Índice 2): TIPO DE MOVIMENTAÇÃO SIADS
      let tipoSiads = String(r[2] || "").trim().toUpperCase();
      tipoSiads = tipoSiads.replace(/[ÁÀÂÃ]/g, "A").replace(/[ÉÈÊ]/g, "E").replace(/[ÍÌÎ]/g, "I").replace(/[ÓÒÔÕ]/g, "O").replace(/[ÚÙÛ]/g, "U");

      if (tipoSiads.indexOf("TRANSFERENCIA ENTRE UG") !== -1) {
        
        // Coluna M (Índice 12): DATA E HORA
        let celulaData = r[12];
        let dataMov = null;

        if (celulaData instanceof Date) {
          dataMov = celulaData;
        } else if (celulaData) {
          dataMov = new Date(celulaData);
        }

        if (dataMov && !isNaN(dataMov.getTime())) {
          if (dataMov.getFullYear() !== anoFiltro || (dataMov.getMonth() + 1) !== mesFiltro) {
            continue;
          }
        } else {
          continue; 
        }

        // Recupera e desmembra os dados das colunas N a S utilizando caractere pipe
        const nlsSaida = String(r[13] || "").split('|');
        const ctsSaida = String(r[14] || "").split('|');
        const vlrsSaida = String(r[15] || "").split('|');
        const nlsEntrada = String(r[16] || "").split('|');
        const ctsEntrada = String(r[17] || "").split('|');
        const vlrsEntrada = String(r[18] || "").split('|');

        const totalSublinhas = Math.max(nlsSaida.length, ctsSaida.length, vlrsSaida.length, nlsEntrada.length, ctsEntrada.length, vlrsEntrada.length, 1);

        for (let s = 0; s < totalSublinhas; s++) {
          
          // 🔥 FUNÇÃO AUXILIAR DE CONVERSÃO ULTRA-SEGURA E CEGA
          function forcarNumeroDecimal(valorRaw) {
            if (!valorRaw) return 0;
            let texto = String(valorRaw).replace(/[R$\s]/g, '').trim();
            if (!texto || texto === "0") return 0;

            // Caso 1: String possui vírgula e ponto (ex: 1.250,82) -> remove ponto e troca vírgula por ponto
            if (texto.includes(",") && texto.includes(".")) {
              texto = texto.replace(/\./g, "").replace(",", ".");
            } 
            // Caso 2: String possui APENAS vírgula (ex: 799,82) -> troca vírgula por ponto
            else if (texto.includes(",") && !texto.includes(".")) {
              texto = texto.replace(",", ".");
            }
            // Caso 3: Já é um padrão puramente decimal (ex: 799.82) -> parseFloat lê direto sem alterações

            let numero = parseFloat(texto);
            return isNaN(numero) ? 0 : numero;
          }

          // Aplica a nova lógica de tratamento individual para saídas e entradas
          let vS = forcarNumeroDecimal(vlrsSaida[s]);
          let vE = forcarNumeroDecimal(vlrsEntrada[s]);
          
          const ugOrigem = String(r[6] || "").trim(); 
          const ugDestino = String(r[7] || "").trim(); 
          const cSaida = String(ctsSaida[s] || "").trim();
          const cEntrada = String(ctsEntrada[s] || "").trim();

          analitico.push({
            sequencial: String(r[0] || "").trim(), 
            processo: String(r[3] || "").trim(),   
            ugOrigem: ugOrigem,
            nlSaida: String(nlsSaida[s] || "").trim(),
            contaSaida: cSaida,
            valorSaida: vS,
            ugDestino: ugDestino,
            nlEntrada: String(nlsEntrada[s] || "").trim(),
            contaEntrada: cEntrada,
            valorEntrada: vE
          });

          // Consolidação por UG/Conta para Saídas
          if (ugOrigem && cSaida) {
            const chaveSaida = ugOrigem + "_" + cSaida;
            if (!mapaResumo[chaveSaida]) {
              mapaResumo[chaveSaida] = { ug: ugOrigem, conta: cSaida, totalEntrada: 0, totalSaida: 0 };
            }
            mapaResumo[chaveSaida].totalSaida += vS;
          }

          // Consolidação por UG/Conta para Entradas
          if (ugDestino && cEntrada) {
            const chaveEntrada = ugDestino + "_" + cEntrada;
            if (!mapaResumo[chaveEntrada]) {
              mapaResumo[chaveEntrada] = { ug: ugDestino, conta: cEntrada, totalEntrada: 0, totalSaida: 0 };
            }
            mapaResumo[chaveEntrada].totalEntrada += vE;
          }
        }
      }
    }

    const resumo = Object.values(mapaResumo).sort((a, b) => {
      if (a.ug !== b.ug) return a.ug.localeCompare(b.ug);
      return a.conta.localeCompare(b.conta);
    });

    return { sucesso: true, analitico: analitico, resumo: resumo };

  } catch (e) {
    return { sucesso: false, message: "Erro no processamento do relatório: " + e.message };
  }
}
