function salvarDados(dados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); 

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaEntrada = ss.getSheetByName("ENTRADA");
    if (!abaEntrada) throw new Error("Aba 'ENTRADA' não encontrada.");

    const dataRegistro = new Date();
    const anoAtual = dataRegistro.getFullYear().toString();
    const usuario = Session.getEffectiveUser().getEmail();

    // --- LÓGICA DO SEQUENCIAL POR PROCESSO ---
    const ultimaLinha = abaEntrada.getLastRow();
    let contadorSequencial = 0;

    if (ultimaLinha > 1) {
      const ultimoValor = abaEntrada.getRange(ultimaLinha, 1).getValue().toString();
      if (ultimoValor.startsWith(anoAtual)) {
        contadorSequencial = parseInt(ultimoValor.substring(4)) || 0;
      }
    }

    // Incrementamos o contador APENAS UMA VEZ para todo este salvamento (um novo processo)
    contadorSequencial++;
    const novoSequencial = anoAtual + contadorSequencial.toString().padStart(6, '0');
    // ------------------------------------------

    const novasLinhas = [];

    // Loop pelas notas e itens
    dados.notas.forEach(nota => {
      nota.itens.forEach(item => {
        
        // Aqui, cada linha recebe o MESMO novoSequencial gerado acima
        const linha = [];
        linha[COL.SEQUENCIAL]           = novoSequencial; // Identificador único do processo
        linha[COL.DATA_REGISTRO]        = dataRegistro;
        linha[COL.USUARIO]              = usuario;
        linha[COL.TIPO]                 = dados.tipo;
        linha[COL.DATA_PROCESSO]        = dados.dataProcesso;
        linha[COL.PROCESSO]             = dados.processo;
        linha[COL.UG]                   = dados.ug;
        linha[COL.UORG]                 = dados.uorg;
        linha[COL.EMPENHO]              = dados.empenho;
        linha[COL.UG_EMPENHO]           = dados.ugEmpenho;
        linha[COL.EDITAL]               = dados.edital;
        linha[COL.BENEFICIARIO]         = dados.beneficiario;
        linha[COL.CPF]                  = dados.cpf;
        linha[COL.SIAPE]                = dados.siape;
        linha[COL.TTP]                  = dados.fade;
        linha[COL.CORRESPONSAVEL]       = dados.corresponsavel;
        linha[COL.CPF_CORRESP]          = dados.cpfCorresponsavel;
        linha[COL.SIAPE_CORRESP]        = dados.siapeCorresponsavel;
        linha[COL.NF_NUMERO]            = nota.numero;
        linha[COL.FORNECEDOR]           = nota.fornecedor;
        linha[COL.VALOR_FRETE_NF]       = nota.valorFrete;
        linha[COL.VALOR_TOTAL_NF]       = nota.valorTotalNota;
        linha[COL.ORDEM_ITEM_NF]        = item.ordem; // Diferencia os itens dentro do mesmo sequencial
        linha[COL.CADMAT]               = item.cadmat;
        linha[COL.INFO_COMPLEMENTAR]    = item.info;
        linha[COL.MARCA]                = item.marca;
        linha[COL.CONTA]                = item.conta;
        linha[COL.QUANTIDADE]           = item.qtd;
        linha[COL.VALOR_UNIT_COM_RATEIO]= item.valorUnitarioRateado;
        linha[COL.TOTAL_ITEM_COM_FRETE] = item.valorTotalItem;

        novasLinhas.push(linha);
      });
    });

    if (novasLinhas.length > 0) {
      abaEntrada.getRange(ultimaLinha + 1, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);
    }

    return { 
      status: "ok", 
      totalItens: novasLinhas.length,
      primeiroSequencial: novoSequencial // Agora é um só para todos
    };

  } catch (e) {
    return { status: "erro", message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Salva os números de tombamento na aba ENTRADA, coluna AE (coluna 31).
 * @param {string} sequencial - O número do registro (Ex: 2026000001).
 * @param {Array} lista - Array de objetos contendo {ordem: "1", valor: "Tombo123"}.
 */
// BUSCA: Traz todos os itens de todas as notas do sequencial
function getItensPorSequencial(sequencial) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ENTRADA");
    const dados = aba.getDataRange().getValues();
    const seqBusca = String(sequencial).replace(/\D/g, '').trim();
    
    const filtrados = dados.slice(1).filter(r => {
      return String(r[COL.SEQUENCIAL]).replace(/\D/g, '').trim() === seqBusca;
    });

    return filtrados.map(r => {
      let vRaw = String(r[COL.VALOR_UNIT_COM_RATEIO] || "0");
      let vLimpo = vRaw.replace(/[R$\s.]/g, '').replace(',', '.');
      let vFinal = parseFloat(vLimpo) || 0;

      return {
        ordem: r[COL.ORDEM_ITEM_NF] || "",
        cadmat: r[COL.CADMAT] || "",
        descricao: r[COL.INFO_COMPLEMENTAR] || "",
        qtd: r[COL.QUANTIDADE] || 0,
        valorUnit: vFinal,
        nota: r[COL.NF_NUMERO] ? String(r[COL.NF_NUMERO]) : "Não Informada"
      };
    });
  } catch (e) {
    console.log("Erro no servidor: " + e.message);
    return [];
  }
}

// SALVAMENTO: Garante que percorre todas as linhas sem parar na primeira nota
// SALVAMENTO: Garante que percorre todas as linhas sem parar na primeira nota e expande intervalos numéricos
function salvarTombamentosNaPlanilha(sequencial, lista) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaEntrada = ss.getSheetByName("ENTRADA");
  const abaTombos = ss.getSheetByName("TOMBOS");
  
  if (!abaEntrada) return "❌ Aba 'ENTRADA' não encontrada.";
  
  const dados = abaEntrada.getDataRange().getValues();
  const COLUNA_AE = 31; // Coluna AE (31) na aba ENTRADA
  let contadorSucesso = 0;

  const seqBusca = String(sequencial).replace(/\D/g, '').trim();

  // Mapeia na ordem física EXATA todas as linhas pertencentes a este Sequencial
  const linhasDoSequencial = [];
  for (let i = 1; i < dados.length; i++) {
    const seqPlanilha = String(dados[i][COL.SEQUENCIAL] || "").replace(/\D/g, '').trim();
    if (seqPlanilha === seqBusca) {
      linhasDoSequencial.push({
        numLinhaPlanilha: i + 1, // Linha real da planilha (1-based)
        dadosLinha: dados[i]
      });
    }
  }

  if (linhasDoSequencial.length === 0) {
    return "❌ Nenhum item encontrado na aba ENTRADA para o sequencial informado.";
  }

  const tombosAdicionadosNesteLote = new Set();

  lista.forEach(itemTombo => {
    const numeroTomboNovo = String(itemTombo.valor).trim();
    const pos = parseInt(itemTombo.posicao, 10);

    if (!numeroTomboNovo || isNaN(pos)) return;

    // Aloca a 1ª caixa (pos 0) na 1ª linha do sequencial, a 2ª (pos 1) na 2ª linha, etc.
    if (pos >= 0 && pos < linhasDoSequencial.length) {
      const alvo = linhasDoSequencial[pos];
      const numLinha = alvo.numLinhaPlanilha;
      const linhaDados = alvo.dadosLinha;

      // Grava no campo correto da aba ENTRADA
      abaEntrada.getRange(numLinha, COLUNA_AE).setValue(numeroTomboNovo);
      contadorSucesso++;

      // Gravação na aba TOMBOS
      if (abaTombos) {
        const ugEntrada   = String(linhaDados[COL.UG] || "").trim();
        const uorgEntrada = String(linhaDados[COL.UORG] || "").trim();
        const cadmat      = String(linhaDados[23] || "").trim(); 
        const infoCompl   = String(linhaDados[24] || "").trim(); 

        const listaTombosParaProcessar = [];
        const regexIntervalo = /^(\d+)\s+a\s+(\d+)$/i;
        const correspondencia = numeroTomboNovo.match(regexIntervalo);

        if (correspondencia) {
          const inicio = parseInt(correspondencia[1], 10);
          const fim = parseInt(correspondencia[2], 10);
          const min = Math.min(inicio, fim);
          const max = Math.max(inicio, fim);

          for (let t = min; t <= max; t++) {
            listaTombosParaProcessar.push(String(t));
          }
        } else {
          listaTombosParaProcessar.push(numeroTomboNovo);
        }

        listaTombosParaProcessar.forEach(tomboIndividual => {
          if (!tombosAdicionadosNesteLote.has(tomboIndividual)) {
            abaTombos.appendRow([
              tomboIndividual, // Coluna A -> tombo_SIADS
              cadmat,          // Coluna B -> CADMAT
              infoCompl,       // Coluna C -> Descrição
              uorgEntrada,     // Coluna D -> uorg
              "",              // Coluna E -> tombo_SIPAC
              ugEntrada        // Coluna F -> UG
            ]);
            tombosAdicionadosNesteLote.add(tomboIndividual);
          }
        });
      }
    }
  });

  return contadorSucesso > 0 
    ? "✅ Sucesso! " + contadorSucesso + " item(ns) atualizado(s) e inserido(s) na aba TOMBOS."
    : "❌ Falha ao vincular os itens.";
}

function getDadosRelatorioEntrada(sequencial) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("ENTRADA");
  if (!aba) return { sucesso: false, message: "Aba ENTRADA não encontrada." };
  
  const dados = aba.getDataRange().getValues();
  const seqBusca = String(sequencial || "").replace(/\D/g, '').trim();
  
  // Filtra todas as linhas pertencentes ao Sequencial usando a constante da coluna
  const linhasFiltradas = [];
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][COL.SEQUENCIAL] || "").replace(/\D/g, '').trim() === seqBusca) {
      linhasFiltradas.push({
        numLinhaPlanilha: i + 1,
        r: dados[i]
      });
    }
  }
  
  if (linhasFiltradas.length === 0) return { sucesso: false };

  const primeira = linhasFiltradas[0].r;
  
  // Captura os códigos originais salvos na planilha via constantes COL
  const codUg = String(primeira[COL.UG] || "").trim();
  const codUorg = String(primeira[COL.UORG] || "").trim();
  const cpfCorresponsavel = String(primeira[COL.CPF_CORRESP] || "").replace(/\D/g, '').trim();

  // --- BUSCA DA DESCRIÇÃO DA UG ---
  let descricaoUg = "";
  const abaUg = ss.getSheetByName("UG");
  if (abaUg && codUg) {
    const dadosUg = abaUg.getDataRange().getValues();
    for (let u = 1; u < dadosUg.length; u++) {
      if (String(dadosUg[u][0] || "").trim() === codUg) {
        const sigla = String(dadosUg[u][2] || "").trim();
        const desc = String(dadosUg[u][1] || "").trim();
        descricaoUg = sigla ? `${desc} (${sigla})` : desc;
        break;
      }
    }
  }

  // --- BUSCA DA DESCRIÇÃO DA UORG ---
  let descricaoUorg = "";
  const abaOrganograma = ss.getSheetByName("ORGANOGRAMA");
  if (abaOrganograma && codUorg) {
    const dadosOrg = abaOrganograma.getDataRange().getValues();
    for (let o = 1; o < dadosOrg.length; o++) {
      if (String(dadosOrg[o][2] || "").trim() === codUorg) { 
        descricaoUorg = String(dadosOrg[o][1] || "").trim(); 
        break;
      }
    }
  }

  // --- BUSCA DOS CONTATOS DO CO-RESPONSÁVEL NA ABA SERVIDORES ---
  let emailCorresponsavel = "Não Cadastrado";
  let telefoneCorresponsavel = "Não Cadastrado";
  const abaServidores = ss.getSheetByName("SERVIDORES");
  if (abaServidores && cpfCorresponsavel) {
    const dadosServidores = abaServidores.getDataRange().getValues();
    for (let s = 1; s < dadosServidores.length; s++) {
      const cpfServidor = String(dadosServidores[s][0] || "").replace(/\D/g, '').trim();
      if (cpfServidor === cpfCorresponsavel) {
        emailCorresponsavel = String(dadosServidores[s][7] || "").trim() || "Não Cadastrado";    
        telefoneCorresponsavel = String(dadosServidores[s][8] || "").trim() || "Não Cadastrado"; 
        break; 
      }
    }
  }

  // Tratamento seguro para formatação de data de registro (Coluna B)
  let dataFormatada = "";
  if (primeira[COL.DOCUMENTODATA_REGISTRO] instanceof Date) {
    dataFormatada = Utilities.formatDate(primeira[COL.DOCUMENTODATA_REGISTRO], Session.getScriptTimeZone(), "dd/MM/yyyy");
  } else {
    dataFormatada = String(primeira[COL.DOCUMENTODATA_REGISTRO] || "").trim();
  }

  // Mapeia os itens baseados estritamente nas constantes originais mapeadas de colunas
  const itensMapeados = linhasFiltradas.map(item => {
    const r = item.r;
    
    let vRaw = String(r[COL.VALOR_UNIT_COM_RATEIO] || "0");
    let vLimpo = vRaw.replace(/[R$\s.]/g, '').replace(',', '.');
    let vFinal = parseFloat(vLimpo) || 0;

    let descCompleta = String(r[COL.INFO_COMPLEMENTAR] || "").trim(); 
    let marca = String(r[COL.MARCA] || "").trim();       
    if (marca) {
      descCompleta += " - Marca: " + marca;
    }

    // Busca o tombamento dinâmico da coluna 31 (Coluna AE)
    let tombamentoReal = String(aba.getRange(item.numLinhaPlanilha, 31).getValue() || "").trim();

    return {
      ordem: String(r[COL.ORDEM_ITEM_NF] || "").trim(),      
      cadmat: String(r[COL.CADMAT] || "").trim(),     
      descricao: descCompleta,                
      qtd: String(r[COL.QUANTIDADE] || "0").trim(),       
      valorUnit: vFinal,                      
      tombamento: tombamentoReal || "Pendente"  
    };
  });

  return {
    sucesso: true,
    dados: {
      sequencial: seqBusca,
      processo: String(primeira[COL.PROCESSO] || "").trim(), 
      data: dataFormatada,                        
      tipo: String(primeira[COL.TIPO] || "").trim(),     
      ug: codUg,
      ugDescricao: descricaoUg, 
      uorg: codUorg,
      uorgDescricao: descricaoUorg, 
      beneficiario: String(primeira[COL.BENEFICIARIO] || "").trim(),
      cpf: String(primeira[COL.CPF] || "").trim(),
      corresponsavel: String(primeira[COL.CORRESPONSAVEL] || "").trim(),
      cpfCorresponsavel: String(primeira[COL.CPF_CORRESP] || "").trim(),
      siapeCorresponsavel: String(primeira[COL.SIAPE_CORRESP] || "").trim(),
      emailCorresponsavel: emailCorresponsavel,
      telefoneCorresponsavel: telefoneCorresponsavel,
      itens: itensMapeados
    }
  };
}

function carregarOrganograma() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ORGANOGRAMA");
    if (!aba) return [];
    
    const dados = aba.getDataRange().getValues();
    // Remove a linha do cabeçalho
    const linhas = dados.slice(1);
    
    // Mapeia as colunas conforme sua especificação:
    // A: Código SIPAC [0], B: Descrição [1], C: UORG [2], D: UG [3]
    return linhas.map(r => ({
      sipac: String(r[0] || "").trim(),
      descricao: String(r[1] || "").trim(),
      uorg: String(r[2] || "").trim(),
      ug: String(r[3] || "").trim()
    })).filter(item => item.ug !== "" && item.uorg !== ""); // Filtra linhas vazias
    
  } catch (e) {
    console.log("Erro ao carregar organograma: " + e.message);
    return [];
  }
}

function carregarListaUGsOrganograma() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("UG");
    if (!aba) return [];
    
    const dados = aba.getDataRange().getValues();
    if (dados.length <= 1) return [];
    
    const linhas = dados.slice(1); // Remove o cabeçalho
    
    // Mapeamento baseado nas colunas: A (UG), B (Descrição), C (Sigla)
    return linhas.map(r => ({
      ug: String(r[0] || "").trim(),
      descricao: String(r[1] || "").trim(),
      sigla: String(r[2] || "").trim()
    })).filter(item => item.ug !== "");
    
  } catch (e) {
    console.log("Erro ao carregar lista de UGs: " + e.message);
    return [];
  }
}

function atualizarContatoServidorGeral(dados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // Evita concorrência na planilha

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("SERVIDORES");
    if (!aba) return { sucesso: false, message: "Aba 'SERVIDORES' não encontrada." };

    const dadosServidores = aba.getDataRange().getValues();
    const cpfBusca = String(dados.cpf || "").replace(/\D/g, '').trim();
    
    if (!cpfBusca) return { sucesso: false, message: "CPF inválido ou não informado." };

    let linhaEncontrada = -1;

    // Procura o CPF na coluna A (índice 0)
    for (let i = 1; i < dadosServidores.length; i++) {
      const cpfServidor = String(dadosServidores[i][0] || "").replace(/\D/g, '').trim();
      if (cpfServidor === cpfBusca) {
        linhaEncontrada = i + 1; // Ajusta o índice para a linha real da planilha
        break;
      }
    }

    if (linhaEncontrada !== -1) {
      // Atualiza o e-mail (Coluna H - coluna 8) e Telefone (Coluna I - coluna 9)
      aba.getRange(linhaEncontrada, 8).setValue(String(dados.email || "").trim());
      aba.getRange(linhaEncontrada, 9).setValue(String(dados.telefone || "").trim());
      return { sucesso: true, message: "Contatos do servidor atualizados com sucesso!" };
    } else {
      // Caso o servidor não exista, cria um novo registro básico
      // Ordem: CPF, Nome, SIAPE, Data Admissão, Tipo, Situação, Órgão, Email, Telefone
      const novaLinha = [
        cpfBusca,
        String(dados.nome || "").trim().toUpperCase(),
        String(dados.siape || "").trim(),
        "", "", "", "", // Colunas intermediárias em branco
        String(dados.email || "").trim(),
        String(dados.telefone || "").trim()
      ];
      aba.appendRow(novaLinha);
      return { sucesso: true, message: "Novo servidor cadastrado com os contatos!" };
    }

  } catch (e) {
    return { sucesso: false, message: "Erro ao salvar contatos: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Busca os dados básicos de um registro patrimonial com base no número Sequencial.
 * Agrupa os itens por Nota Fiscal para que o usuário informe uma NL por NF.
 * @param {string} sequencialBusca - O número sequencial digitado (ex: 2026000001)
 */
function buscarRegistroParaNL(sequencialBusca) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ENTRADA");
    if (!aba) return { sucesso: false, message: "Aba 'ENTRADA' não encontrada." };

    const dados = aba.getDataRange().getValues();
    if (dados.length <= 1) return { sucesso: false, message: "Nenhum registro cadastrado no sistema." };

    const sequencialLimpo = String(sequencialBusca).trim();
    const notasLocalizadas = {};
    let achou = false;

    // Percorre a aba ENTRADA procurando o sequencial
    for (let i = 1; i < dados.length; i++) {
      const r = dados[i];
      const seqPlanilha = String(r[COL.SEQUENCIAL] || "").trim();

      if (seqPlanilha === sequencialLimpo) {
        achou = true;
        const nfNumero = String(r[COL.NF_NUMERO] || "Não Informada").trim();

        // Tratamento do valor monetário do item
        let valorBruto = r[COL.TOTAL_ITEM_COM_FRETE];
        let valorItem = 0;
        if (typeof valorBruto === "number") {
          valorItem = valorBruto;
        } else if (valorBruto) {
          let vStr = String(valorBruto).trim().replace(/R\$/g, "").replace(/\s/g, "");
          if (vStr.includes(",") && vStr.includes(".")) vStr = vStr.replace(/\./g, "").replace(",", ".");
          else if (vStr.includes(",")) vStr = vStr.replace(",", ".");
          valorItem = parseFloat(vStr) || 0;
        }

        // Se a nota fiscal ainda não foi catalogada neste agrupamento, inicializa os dados básicos
        if (!notasLocalizadas[nfNumero]) {
          notasLocalizadas[nfNumero] = {
            sequencial: seqPlanilha,
            tipoEntrada: String(r[COL.TIPO] || "NÃO ESPECIFICADO").trim(),
            processo: String(r[COL.PROCESSO] || "-").trim(),
            ug: String(r[COL.UG] || "-").trim(),
            documentoNF: nfNumero,
            valorTotalNF: 0,
            nlAtual: String(r[31] || "").trim(), // Coluna AF é o índice 31 (base 0)
            linhasPlanilha: [] // Guarda os índices reais da planilha para a hora do salvamento
          };
        }

        // Acumula o valor total da Nota Fiscal (Soma de todos os seus itens)
        notasLocalizadas[nfNumero].valorTotalNF = Math.round((notasLocalizadas[nfNumero].valorTotalNF + valorItem) * 100) / 100;
        
        // Armazena o número da linha física (i + 1) para atualização direta posterior
        notasLocalizadas[nfNumero].linhasPlanilha.push(i + 1);
      }
    }

    if (!achou) {
      return { sucesso: false, message: "Sequencial de registro não localizado na base de dados." };
    }

    return { sucesso: true, notas: Object.values(notasLocalizadas) };

  } catch (e) {
    console.log("Erro ao buscar sequencial para NL: " + e.message);
    return { sucesso: false, message: e.message };
  }
}

/**
 * Grava o número da NL informado pelo usuário na coluna AF (coluna 32) de todas as linhas associadas.
 * @param {Array} dadosSalvar - Vetor de objetos contendo { linhas: [3, 4], nl: "2026NL00123" }
 */
function salvarNLRegistro(dadosSalvar) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("ENTRADA");
    if (!aba) return { sucesso: false, message: "Aba 'ENTRADA' não encontrada." };

    // Percorre cada bloco de Nota Fiscal enviado pelo formulário
    dadosSalvar.forEach(bloco => {
      const numeroNL = String(bloco.nl || "").trim().toUpperCase();
      
      // Atualiza a coluna AF (coluna 32) para cada linha do item daquela NF
      bloco.linhas.forEach(linhaReal => {
        aba.getRange(linhaReal, 32).setValue(numeroNL); 
      });
    });

    return { sucesso: true, message: "Número(s) de NL atualizado(s) com sucesso na base de dados!" };

  } catch (e) {
    console.log("Erro ao salvar NL: " + e.message);
    return { sucesso: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}