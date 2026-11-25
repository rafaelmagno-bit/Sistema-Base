// ============================================================================
// ⚙️ MAPEAMENTO DE COLUNAS DA ABA "ENTRADA" (Índice 0 = Coluna A)
// Ajuste os índices abaixo conforme a estrutura exata da sua planilha
// ============================================================================
const COL_ENTRADA = {
  SEQUENCIAL: 0,    // Coluna A
  DATA_REGISTRO: 1, // Coluna B
  USUARIO: 2,       // Coluna C (Ajuste se o e-mail/usuário estiver em outra coluna)
  TIPO: 3,          // Coluna D
  DATA_DOC: 4,      // Coluna E
  PROCESSO: 5,      // Coluna F
  UG: 6,            // Coluna G
  UORG: 7,          // Coluna H
  VALOR: 8,         // Coluna I (Ajuste conforme sua planilha)
  STATUS: 15        // Coluna P (Sugestão: coluna para marcar "ATIVO" ou "ANULADO")
};

const MSG_APENAS_ADMIN = "Esta ação somente poderá ser realizada por usuários com o perfil ADMIN.";

/**
 * Busca entradas na aba "ENTRADA" aplicando os filtros informados.
 * Agrupa múltiplas linhas de itens pertencentes ao mesmo Sequencial.
 */
function consultarEntradasFiltradas(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ENTRADA");
    if (!aba) throw new Error("Aba 'ENTRADA' não encontrada.");

    const dados = aba.getDataRange().getValues();
    if (dados.length <= 1) return { sucesso: true, registros: [], perfil: getPerfilUsuario() };

    const dtInicio = filtros.dataInicio ? new Date(filtros.dataInicio + "T00:00:00") : null;
    const dtFim = filtros.dataFim ? new Date(filtros.dataFim + "T23:59:59") : null;
    const filtroUsuario = String(filtros.usuario || "").trim().toLowerCase();
    const filtroSeq = String(filtros.sequencial || "").replace(/\D/g, "").trim();
    const filtroProc = String(filtros.processo || "").trim().toLowerCase();

    const agrupadoPorSequencial = {};

    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      const seq = String(linha[COL_ENTRADA.SEQUENCIAL] || "").trim();
      if (!seq) continue;

      // 1. Filtro por Sequencial
      if (filtroSeq && !seq.replace(/\D/g, "").includes(filtroSeq)) continue;

      // 2. Filtro por Processo
      const proc = String(linha[COL_ENTRADA.PROCESSO] || "").trim();
      if (filtroProc && !proc.toLowerCase().includes(filtroProc)) continue;

      // 3. Filtro por Usuário
      const usr = String(linha[COL_ENTRADA.USUARIO] || "").trim();
      if (filtroUsuario && !usr.toLowerCase().includes(filtroUsuario)) continue;

      // 4. Filtro por Data de Registro (De / Até)
      const rawDataReg = linha[COL_ENTRADA.DATA_REGISTRO];
      const dataReg = (rawDataReg instanceof Date) ? rawDataReg : new Date(rawDataReg);
      if (dtInicio && (!isNaN(dataReg) && dataReg < dtInicio)) continue;
      if (dtFim && (!isNaN(dataReg) && dataReg > dtFim)) continue;

      if (!agrupadoPorSequencial[seq]) {
        agrupadoPorSequencial[seq] = {
          sequencial: seq,
          dataRegistroISO: !isNaN(dataReg) ? Utilities.formatDate(dataReg, Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
          dataRegistroFmt: !isNaN(dataReg) ? Utilities.formatDate(dataReg, Session.getScriptTimeZone(), "dd/MM/yyyy") : "-",
          mesRegistro: !isNaN(dataReg) ? dataReg.getMonth() : -1,
          anoRegistro: !isNaN(dataReg) ? dataReg.getFullYear() : -1,
          usuario: usr,
          tipo: String(linha[COL_ENTRADA.TIPO] || ""),
          dataDocISO: formatarParaInputDate_(linha[COL_ENTRADA.DATA_DOC]),
          processo: proc,
          ug: String(linha[COL_ENTRADA.UG] || ""),
          uorg: String(linha[COL_ENTRADA.UORG] || ""),
          status: String(linha[COL_ENTRADA.STATUS] || "ATIVO").toUpperCase(),
          qtdItens: 0
        };
      }
      agrupadoPorSequencial[seq].qtdItens++;
    }

    return {
      sucesso: true,
      perfil: getPerfilUsuario(),
      registros: Object.values(agrupadoPorSequencial).reverse() // Mais recentes primeiro
    };
  } catch (e) {
    return { sucesso: false, message: e.message };
  }
}

/**
 * Altera os dados de uma entrada (todas as linhas vinculadas ao sequencial).
 * Regra: Mês corrente -> Qualquer operador autorizado | Meses anteriores -> Apenas ADMIN.
 */
function alterarEntradaRegistro(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const perfil = getPerfilUsuario();
    if (perfil === "VISUALIZADOR") {
      throw new Error("Seu perfil (VISUALIZADOR) não permite alterar registros.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ENTRADA");
    const dados = aba.getDataRange().getValues();

    const hoje = new Date();
    const mesAtual = hoje.getMonth();
    const anoAtual = hoje.getFullYear();

    let linhasAtualizadas = 0;

    for (let i = 1; i < dados.length; i++) {
      const seqLinha = String(dados[i][COL_ENTRADA.SEQUENCIAL] || "").trim();
      if (seqLinha === String(payload.sequencial).trim()) {
        
        // Verifica se a entrada está anulada
        const statusAtual = String(dados[i][COL_ENTRADA.STATUS] || "ATIVO").toUpperCase();
        if (statusAtual === "ANULADO") {
          throw new Error("Não é possível alterar uma entrada que já foi anulada.");
        }

        // Valida regra do Mês Corrente vs Mês Anterior
        const rawDataReg = dados[i][COL_ENTRADA.DATA_REGISTRO];
        const dataReg = (rawDataReg instanceof Date) ? rawDataReg : new Date(rawDataReg);
        const ehMesCorrente = (dataReg.getMonth() === mesAtual && dataReg.getFullYear() === anoAtual);

        if (!ehMesCorrente && perfil !== "ADMIN") {
          throw new Error(MSG_APENAS_ADMIN);
        }

        // Atualiza os campos editáveis nas colunas correspondentes (índice + 1)
        const numLinha = i + 1;
        aba.getRange(numLinha, COL_ENTRADA.PROCESSO + 1).setValue(payload.processo);
        aba.getRange(numLinha, COL_ENTRADA.TIPO + 1).setValue(payload.tipo);
        aba.getRange(numLinha, COL_ENTRADA.UG + 1).setValue(payload.ug);
        aba.getRange(numLinha, COL_ENTRADA.UORG + 1).setValue(payload.uorg);
        if (payload.dataDoc) {
          aba.getRange(numLinha, COL_ENTRADA.DATA_DOC + 1).setValue(new Date(payload.dataDoc + "T12:00:00"));
        }

        linhasAtualizadas++;
      }
    }

    if (linhasAtualizadas === 0) throw new Error("Sequencial não localizado para alteração.");

    return { sucesso: true, message: "Entrada atualizada com sucesso!" };
  } catch (e) {
    return { sucesso: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Anula uma entrada marcando seu status como "ANULADO" (Preserva trilha de auditoria).
 * Regra: Exclusivo para perfil ADMIN.
 */
function anularEntradaRegistro(sequencial, motivo) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    // Bloqueio rigoroso no servidor
    if (getPerfilUsuario() !== "ADMIN") {
      throw new Error(MSG_APENAS_ADMIN);
    }

    if (!motivo || String(motivo).trim().length < 5) {
      throw new Error("Informe uma justificativa válida (mínimo 5 caracteres) para anular.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ENTRADA");
    const dados = aba.getDataRange().getValues();
    const usuarioAdmin = Session.getEffectiveUser().getEmail();
    const dataAnulacao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

    let encontrou = false;
    const textoStatus = `ANULADO (${usuarioAdmin} em ${dataAnulacao} - Motivo: ${motivo.trim()})`;

    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][COL_ENTRADA.SEQUENCIAL] || "").trim() === String(sequencial).trim()) {
        aba.getRange(i + 1, COL_ENTRADA.STATUS + 1).setValue(textoStatus);
        encontrou = true;
      }
    }

    if (!encontrou) throw new Error("Sequencial não encontrado.");

    return { sucesso: true, message: "Entrada anulada com sucesso." };
  } catch (e) {
    return { sucesso: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function formatarParaInputDate_(valor) {
  if (!valor) return "";
  const d = (valor instanceof Date) ? valor : new Date(valor);
  if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}