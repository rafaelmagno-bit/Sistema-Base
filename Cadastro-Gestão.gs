// ============================================================================
// ⚙️ MAPEAMENTO EXATO DAS COLUNAS DA ABA "ENTRADA" (Índice 0 = Coluna A)
// ============================================================================
const COL_ENTRADA = {
  SEQUENCIAL: 0,         // A  - SEQUENCIAL DOCUMENTO
  DATA_REGISTRO: 1,      // B  - DATA_REGISTRO
  USUARIO: 2,            // C  - USUARIO
  TIPO: 3,               // D  - TIPO
  DATA_PROCESSO: 4,      // E  - DATA DO PROCESSO
  PROCESSO: 5,           // F  - PROCESSO
  UG: 6,                 // G  - UG
  UORG: 7,               // H  - UORG
  EMPENHO: 8,            // I  - EMPENHO
  UG_EMPENHO: 9,         // J  - UG EMPENHO
  EDITAL: 10,            // K  - EDITAL
  BENEFICIARIO: 11,      // L  - BENEFICIARIO
  CPF: 12,               // M  - CPF
  SIAPE: 13,             // N  - SIAPE
  TTP: 14,               // O  - TTP
  CORRESPONSAVEL: 15,    // P  - COORESPONSÁVEL
  CPF_CORRESP: 16,       // Q  - CPF CORRESPONSÁVEL
  SIAPE_CORRESP: 17,     // R  - SIAPE CORRESPONSÁVEL
  NF_NUMERO: 18,         // S  - Nº DOCUMENTO
  FORNECEDOR: 19,        // T  - FORNECEDOR
  VLR_FRETE: 20,         // U  - VLR. FRETE
  VLR_TOTAL_NF: 21,      // V  - VLR TOTAL NF
  ORDEM_ITEM: 22,        // W  - ORDEM DOS ITENS DA NF
  CADMAT: 23,            // X  - CADMAT
  DESCRICAO: 24,         // Y  - INFORMAÇÃO COMPLEMENTAR
  MARCA: 25,             // Z  - MARCA
  CONTA: 26,             // AA - CONTA
  QUANTIDADE: 27,        // AB - QUANTIDADE
  VLR_UNITARIO: 28,      // AC - VALOR UNITARIO
  TOTAL_ITEM: 29,        // AD - TOTAL DO ITEM
  TOMBAMENTO: 30,        // AE - TOMBAMENTO
  NL: 31,                // AF - NL
  STATUS: 32,            // AG - STATUS ("ATIVO" ou "ANULADO")
  USUARIO_ANULACAO: 33,  // AH - USUARIO DA ANULAÇÃO
  DATA_ANULACAO: 34,      // AI - DATA DA ANULAÇÃO
  MOTIVO_ANULACAO: 35    // AJ - MOTIVO DA ANULAÇÃO
};

const MSG_APENAS_ADMIN = "Esta ação somente poderá ser realizada por usuários com o perfil ADMIN.";

/**
 * Converte data da planilha (seja Date, "dd/MM/yyyy" ou "yyyy-MM-dd") para objeto Date válido.
 */
function parseDataPlanilha_(valor) {
  if (!valor) return null;
  if (valor instanceof Date && !isNaN(valor.getTime())) return valor;

  const str = String(valor).trim();
  // Formato dd/MM/yyyy
  if (/^\d{2}\/\d{2}\/\d{4}/.test(str)) {
    const partes = str.split(" ")[0].split("/");
    return new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]), 12, 0, 0);
  }
  // Formato yyyy-MM-dd
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    const partes = str.split("T")[0].split("-");
    return new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]), 12, 0, 0);
  }
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Consulta entradas na aba "ENTRADA" aplicando filtros.
 * Inicia em i = 2 porque as linhas 1 e 2 da planilha são cabeçalhos.
 */
function consultarEntradasFiltradas(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ENTRADA");
    if (!aba) throw new Error("Aba 'ENTRADA' não encontrada.");

    const dados = aba.getDataRange().getValues();
    if (dados.length <= 2) return { sucesso: true, registros: [], perfil: getPerfilUsuario() };

    const dtInicio = filtros.dataInicio ? new Date(filtros.dataInicio + "T00:00:00") : null;
    const dtFim = filtros.dataFim ? new Date(filtros.dataFim + "T23:59:59") : null;
    const filtroUsuario = String(filtros.usuario || "").trim().toLowerCase();
    const filtroSeq = String(filtros.sequencial || "").replace(/\D/g, "").trim();
    const filtroProc = String(filtros.processo || "").trim().toLowerCase();

    const agrupadoPorSequencial = {};
    const tz = Session.getScriptTimeZone();

    // i = 2 pula as duas linhas de cabeçalho da aba ENTRADA
    for (let i = 2; i < dados.length; i++) {
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
      const dataReg = parseDataPlanilha_(linha[COL_ENTRADA.DATA_REGISTRO]);
      if (dtInicio && dataReg && dataReg < dtInicio) continue;
      if (dtFim && dataReg && dataReg > dtFim) continue;

      if (!agrupadoPorSequencial[seq]) {
        const dataProc = parseDataPlanilha_(linha[COL_ENTRADA.DATA_PROCESSO]);
        const statusBruto = String(linha[COL_ENTRADA.STATUS] || "ATIVO").trim().toUpperCase();
        const usrAnul = String(linha[COL_ENTRADA.USUARIO_ANULACAO] || "").trim();
        const motivoAnul = String(linha[COL_ENTRADA.MOTIVO_ANULACAO] || "").trim();
        const dtAnul = linha[COL_ENTRADA.DATA_ANULACAO]
          ? Utilities.formatDate(parseDataPlanilha_(linha[COL_ENTRADA.DATA_ANULACAO]) || new Date(), tz, "dd/MM/yyyy HH:mm")
          : "";

        agrupadoPorSequencial[seq] = {
          sequencial: seq,
          dataRegistroFmt: dataReg ? Utilities.formatDate(dataReg, tz, "dd/MM/yyyy") : "-",
          mesRegistro: dataReg ? dataReg.getMonth() : -1,
          anoRegistro: dataReg ? dataReg.getFullYear() : -1,
          usuario: usr,
          tipo: String(linha[COL_ENTRADA.TIPO] || ""),
          dataDocISO: dataProc ? Utilities.formatDate(dataProc, tz, "yyyy-MM-dd") : "",
          processo: proc,
          ug: String(linha[COL_ENTRADA.UG] || ""),
          uorg: String(linha[COL_ENTRADA.UORG] || ""),
          nl: String(linha[COL_ENTRADA.NL] || ""),
          tombamento: String(linha[COL_ENTRADA.TOMBAMENTO] || ""),
          status: statusBruto,
          motivoAnulacao: motivoAnul,
          infoAnulacao: statusBruto === "ANULADO" ? `Anulado por ${usrAnul} em ${dtAnul}${motivoAnul ? " | Motivo: " + motivoAnul : ""}`
            : "",
          qtdItens: 0
        };
      }
      agrupadoPorSequencial[seq].qtdItens++;
    }

    return {
      sucesso: true,
      perfil: getPerfilUsuario(),
      registros: Object.values(agrupadoPorSequencial).reverse()
    };
  } catch (e) {
    return { sucesso: false, message: e.message };
  }
}

/**
 * Altera os dados de cabeçalho de uma entrada em todas as linhas do sequencial.
 * Mês corrente -> Qualquer operador | Meses anteriores -> Apenas ADMIN.
 */
function alterarEntradaRegistro(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const perfil = getPerfilUsuario();
    if (perfil === "VISUALIZADOR") {
      throw new Error(MSG_APENAS_ADMIN);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ENTRADA");
    const dados = aba.getDataRange().getValues();

    const hoje = new Date();
    const mesAtual = hoje.getMonth();
    const anoAtual = hoje.getFullYear();

    let linhasAtualizadas = 0;

    for (let i = 2; i < dados.length; i++) {
      const seqLinha = String(dados[i][COL_ENTRADA.SEQUENCIAL] || "").trim();
      if (seqLinha === String(payload.sequencial).trim()) {

        const statusAtual = String(dados[i][COL_ENTRADA.STATUS] || "ATIVO").trim().toUpperCase();
        if (statusAtual === "ANULADO") {
          throw new Error("Não é possível alterar uma entrada que já foi anulada.");
        }

        const dataReg = parseDataPlanilha_(dados[i][COL_ENTRADA.DATA_REGISTRO]);
        const ehMesCorrente = dataReg && (dataReg.getMonth() === mesAtual && dataReg.getFullYear() === anoAtual);

        if (!ehMesCorrente && perfil !== "ADMIN") {
          throw new Error(MSG_APENAS_ADMIN);
        }

        const numLinha = i + 1;
        aba.getRange(numLinha, COL_ENTRADA.PROCESSO + 1).setValue(payload.processo);
        aba.getRange(numLinha, COL_ENTRADA.TIPO + 1).setValue(payload.tipo);
        aba.getRange(numLinha, COL_ENTRADA.UG + 1).setValue(payload.ug);
        aba.getRange(numLinha, COL_ENTRADA.UORG + 1).setValue(payload.uorg);
        if (payload.dataDoc) {
          aba.getRange(numLinha, COL_ENTRADA.DATA_PROCESSO + 1).setValue(payload.dataDoc); // Mantém yyyy-MM-dd da planilha
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
 * Anula uma entrada preenchendo:
 * - Coluna AG (STATUS): "ANULADO"
 * - Coluna AH (USUARIO DA ANULAÇÃO): e-mail do ADMIN + motivo
 * - Coluna AI (DATA DA ANULAÇÃO): data/hora atual
 */
function anularEntradaRegistro(sequencial, motivo) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    if (getPerfilUsuario() !== "ADMIN") {
      throw new Error(MSG_APENAS_ADMIN);
    }

    if (!motivo || String(motivo).trim().length < 5) {
      throw new Error("Informe uma justificativa válida (mínimo 5 caracteres) para anular.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ENTRADA");
    const dados = aba.getDataRange().getValues();
    const usuarioAdmin = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    const agora = new Date();

    let encontrou = false;

    for (let i = 2; i < dados.length; i++) {
      if (String(dados[i][COL_ENTRADA.SEQUENCIAL] || "").trim() === String(sequencial).trim()) {
        const numLinha = i + 1;
        aba.getRange(numLinha, COL_ENTRADA.STATUS + 1).setValue("ANULADO");                        // Coluna AG (33)
        aba.getRange(numLinha, COL_ENTRADA.USUARIO_ANULACAO + 1).setValue(usuarioAdmin);          // Coluna AH (34)
        aba.getRange(numLinha, COL_ENTRADA.DATA_ANULACAO + 1).setValue(agora);                    // Coluna AI (35)
        aba.getRange(numLinha, COL_ENTRADA.MOTIVO_ANULACAO + 1).setValue(motivo.trim());          // Coluna AJ (36)
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

/**
 * Carrega todos os dados de um sequencial (cabeçalho, NFs e itens) para edição completa.
 */
function obterEntradaCompletaParaEdicao(sequencial) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ENTRADA");
    const dados = aba.getDataRange().getValues();
    const tz = Session.getScriptTimeZone();

    const linhasSeq = [];
    for (let i = 2; i < dados.length; i++) {
      if (String(dados[i][COL_ENTRADA.SEQUENCIAL] || "").trim() === String(sequencial).trim()) {
        linhasSeq.push(dados[i]);
      }
    }

    if (linhasSeq.length === 0) throw new Error("Sequencial não encontrado.");

    const primeira = linhasSeq[0];
    const status = String(primeira[COL_ENTRADA.STATUS] || "ATIVO").trim().toUpperCase();
    if (status === "ANULADO") throw new Error("Não é possível alterar uma entrada anulada.");

    // Validação de permissão (Mês Corrente vs ADMIN)
    const perfil = getPerfilUsuario();
    const dataReg = parseDataPlanilha_(primeira[COL_ENTRADA.DATA_REGISTRO]);
    const hoje = new Date();
    const ehMesCorrente = dataReg && (dataReg.getMonth() === hoje.getMonth() && dataReg.getFullYear() === hoje.getFullYear());

    if (perfil === "VISUALIZADOR" || (!ehMesCorrente && perfil !== "ADMIN")) {
      throw new Error(MSG_APENAS_ADMIN);
    }

    const dataProc = parseDataPlanilha_(primeira[COL_ENTRADA.DATA_PROCESSO]);

    // Agrupa os itens por Nota Fiscal (Nº Documento)
    const mapaNFs = {};
    linhasSeq.forEach(l => {
      const numNF = String(l[COL_ENTRADA.NF_NUMERO] || "SEM_NF").trim();
      if (!mapaNFs[numNF]) {
        mapaNFs[numNF] = {
          numero: l[COL_ENTRADA.NF_NUMERO] || "",
          fornecedor: l[COL_ENTRADA.FORNECEDOR] || "",
          frete: Number(l[COL_ENTRADA.VLR_FRETE]) || 0,
          totalNF: Number(l[COL_ENTRADA.VLR_TOTAL_NF]) || 0,
          itens: []
        };
      }
      mapaNFs[numNF].itens.push({
        ordem: l[COL_ENTRADA.ORDEM_ITEM] || 1,
        cadmat: String(l[COL_ENTRADA.CADMAT] || "").trim(),
        descricao: String(l[COL_ENTRADA.DESCRICAO] || "").trim(),
        marca: String(l[COL_ENTRADA.MARCA] || "").trim(),
        conta: String(l[COL_ENTRADA.CONTA] || "").trim(),
        qtd: Number(l[COL_ENTRADA.QUANTIDADE]) || 1,
        valorUnit: Number(l[COL_ENTRADA.VLR_UNITARIO]) || 0,
        totalItem: Number(l[COL_ENTRADA.TOTAL_ITEM]) || 0,
        tombamento: String(l[COL_ENTRADA.TOMBAMENTO] || "").trim(),
        nl: String(l[COL_ENTRADA.NL] || "").trim()
      });
    });

    return {
      sucesso: true,
      dados: {
        sequencial: String(primeira[COL_ENTRADA.SEQUENCIAL]).trim(),
        tipo: String(primeira[COL_ENTRADA.TIPO] || "").trim(),
        processo: String(primeira[COL_ENTRADA.PROCESSO] || "").trim(),
        dataProcesso: dataProc ? Utilities.formatDate(dataProc, tz, "yyyy-MM-dd") : "",
        ug: String(primeira[COL_ENTRADA.UG] || "").trim(),
        uorg: String(primeira[COL_ENTRADA.UORG] || "").trim(),
        empenho: String(primeira[COL_ENTRADA.EMPENHO] || "").trim(),
        ugEmpenho: String(primeira[COL_ENTRADA.UG_EMPENHO] || "").trim(),
        edital: String(primeira[COL_ENTRADA.EDITAL] || "").trim(),
        beneficiario: String(primeira[COL_ENTRADA.BENEFICIARIO] || "").trim(),
        cpf: String(primeira[COL_ENTRADA.CPF] || "").trim(),
        siape: String(primeira[COL_ENTRADA.SIAPE] || "").trim(),
        fade: String(primeira[COL_ENTRADA.TTP] || "").trim(),
        corresponsavel: String(primeira[COL_ENTRADA.CORRESPONSAVEL] || "").trim(),
        cpfCorresp: String(primeira[COL_ENTRADA.CPF_CORRESP] || "").trim(),
        siapeCorresp: String(primeira[COL_ENTRADA.SIAPE_CORRESP] || "").trim(),
        notasFiscais: Object.values(mapaNFs)
      }
    };
  } catch (e) {
    return { sucesso: false, message: e.message };
  }
}

/**
 * Salva a edição completa do formulário de cadastro mantendo Sequencial, Data Registro e Tombos/NL.
 */
function salvarEdicaoCompletaEntrada(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const perfil = getPerfilUsuario();
    if (perfil === "VISUALIZADOR") throw new Error(MSG_APENAS_ADMIN);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ENTRADA");
    const dados = aba.getDataRange().getValues();
    const seqAlvo = String(payload.sequencial).trim();

    // Localiza as linhas atuais desse sequencial para preservar metadados (Data Registro, Usuário, Tombos, NL)
    const indicesLinhasAntigas = [];
    let dataRegistroOriginal = new Date();
    let usuarioOriginal = "";
    const mapaPreservados = {}; // Chave: "ordem" -> { tombamento, nl }

    for (let i = 2; i < dados.length; i++) {
      if (String(dados[i][COL_ENTRADA.SEQUENCIAL] || "").trim() === seqAlvo) {
        if (indicesLinhasAntigas.length === 0) {
          dataRegistroOriginal = dados[i][COL_ENTRADA.DATA_REGISTRO];
          usuarioOriginal = dados[i][COL_ENTRADA.USUARIO];

          const dtReg = parseDataPlanilha_(dataRegistroOriginal);
          const hoje = new Date();
          const ehMesCorrente = dtReg && (dtReg.getMonth() === hoje.getMonth() && dtReg.getFullYear() === hoje.getFullYear());
          if (!ehMesCorrente && perfil !== "ADMIN") {
            throw new Error(MSG_APENAS_ADMIN);
          }
        }
        const ordem = String(dados[i][COL_ENTRADA.ORDEM_ITEM] || "").trim();
        mapaPreservados[ordem] = {
          tombamento: dados[i][COL_ENTRADA.TOMBAMENTO] || "",
          nl: dados[i][COL_ENTRADA.NL] || ""
        };
        indicesLinhasAntigas.push(i + 1); // Número da linha 1-based
      }
    }

    if (indicesLinhasAntigas.length === 0) {
      throw new Error("Sequencial original não encontrado na planilha.");
    }

    // Monta as novas linhas completas (36 colunas: A até AJ)
    const novasLinhas = [];
    let contadorOrdem = 1;

    payload.notasFiscais.forEach(nf => {
      nf.itens.forEach(item => {
        const ordStr = String(contadorOrdem);
        const pres = mapaPreservados[ordStr] || { tombamento: "", nl: "" };

        novasLinhas.push([
          seqAlvo,                        // A  - SEQUENCIAL
          dataRegistroOriginal,           // B  - DATA_REGISTRO (preservada)
          usuarioOriginal,                // C  - USUARIO (preservado)
          payload.tipo,                   // D  - TIPO
          payload.dataProcesso,           // E  - DATA DO PROCESSO
          payload.processo,               // F  - PROCESSO
          payload.ug,                     // G  - UG
          payload.uorg,                   // H  - UORG
          payload.empenho || "",          // I  - EMPENHO
          payload.ugEmpenho || "",        // J  - UG EMPENHO
          payload.edital || "",           // K  - EDITAL
          payload.beneficiario || "",     // L  - BENEFICIARIO
          payload.cpf || "",              // M  - CPF
          payload.siape || "",            // N  - SIAPE
          payload.fade || "",             // O  - TTP
          payload.corresponsavel || "",   // P  - COORESPONSÁVEL
          payload.cpfCorresp || "",       // Q  - CPF CORRESPONSÁVEL
          payload.siapeCorresp || "",     // R  - SIAPE CORRESPONSÁVEL
          nf.numero || "",                // S  - Nº DOCUMENTO
          nf.fornecedor || "",            // T  - FORNECEDOR
          nf.frete || 0,                  // U  - VLR. FRETE
          nf.totalNF || 0,                // V  - VLR TOTAL NF
          contadorOrdem,                  // W  - ORDEM DOS ITENS DA NF
          item.cadmat || "",              // X  - CADMAT
          item.descricao || "",           // Y  - INFORMAÇÃO COMPLEMENTAR
          item.marca || "",               // Z  - MARCA
          item.conta || "",               // AA - CONTA
          item.qtd || 1,                  // AB - QUANTIDADE
          item.valorUnit || 0,            // AC - VALOR UNITARIO
          item.totalItem || 0,            // AD - TOTAL DO ITEM
          pres.tombamento,                // AE - TOMBAMENTO (preservado)
          pres.nl,                        // AF - NL (preservado)
          "ATIVO",                        // AG - STATUS
          "",                             // AH - USUARIO DA ANULAÇÃO
          "",                             // AI - DATA DA ANULAÇÃO
          ""                              // AJ - MOTIVO DA ANULAÇÃO
        ]);
        contadorOrdem++;
      });
    });

    // Remove as linhas antigas de baixo para cima e insere as atualizadas
    indicesLinhasAntigas.reverse().forEach(numLinha => aba.deleteRow(numLinha));
    const linhaInsercao = indicesLinhasAntigas[indicesLinhasAntigas.length - 1];
    aba.insertRowsBefore(linhaInsercao, novasLinhas.length);
    aba.getRange(linhaInsercao, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);

    return { status: "ok", message: `Entrada ${seqAlvo} atualizada com sucesso!` };
  } catch (e) {
    return { status: "erro", message: e.message };
  } finally {
    lock.releaseLock();
  }
}