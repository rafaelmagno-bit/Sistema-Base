/**
 * Processa a planilha de Razão de Bens Móveis enviada pelo frontend
 * @param {string} base64Data Conteúdo do arquivo em formato Base64
 * @param {string} fileName Nome do arquivo enviado
 */
function processarRazaoTransferencias(base64Data, fileName) {
  try {
    if (typeof Drive === 'undefined' || !Drive.Files) {
      throw new Error("O serviço avançado 'Drive API' não está ativado. Clique no ícone de (+) em 'Serviços' no Apps Script e adicione a 'Drive API'.");
    }

    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName);
    
    var tempFileName = "TEMP_TRANSF_" + new Date().getTime();
    var tempFile;

    if (typeof Drive.Files.create === 'function') {
      tempFile = Drive.Files.create({ name: tempFileName, mimeType: MimeType.GOOGLE_SHEETS }, blob);
    } else if (typeof Drive.Files.insert === 'function') {
      tempFile = Drive.Files.insert({ title: tempFileName, mimeType: MimeType.GOOGLE_SHEETS }, blob);
    } else {
      throw new Error("Não foi possível acessar os métodos da Drive API.");
    }

    var ss = SpreadsheetApp.openById(tempFile.id);
    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    
    // Remove o arquivo temporário criado no Drive
    DriveApp.getFileById(tempFile.id).setTrashed(true);

    if (data.length < 2) {
      return { sucesso: true, dados: [] };
    }

    // Localização dinâmica da linha de cabeçalho
    var headerRowIndex = -1;
    for (var r = 0; r < Math.min(10, data.length); r++) {
      var rowStr = data[r].join(" ").toLowerCase();
      if (rowStr.indexOf("conta contábil") !== -1 || rowStr.indexOf("documento") !== -1) {
        headerRowIndex = r;
        break;
      }
    }

    if (headerRowIndex === -1) {
      throw new Error("Não foi possível identificar a linha do cabeçalho na planilha.");
    }

    var headers = data[headerRowIndex].map(function(h) { return String(h).trim().toLowerCase(); });
    
    var idxConta = headers.indexOf("conta contábil");
    var idxUg = headers.indexOf("ug executora");
    
    // Identificação da Coluna F (índice 5) para a NL
    var idxDoc = headers.indexOf("documento");
    if (idxDoc === -1) {
      idxDoc = 5; // Coluna F (índice 5 no vetor de dados)
    }

    var idxObs = -1;
    for (var c = 0; c < headers.length; c++) {
      if (headers[c].indexOf("observação") !== -1 || headers[c].indexOf("observacao") !== -1 || headers[c].indexOf("doc - observação") !== -1) {
        idxObs = c;
        break;
      }
    }
    if (idxObs === -1 && headers.length > 7) {
      idxObs = 7; // Coluna H
    }
    
    var idxMovDev = -1;
    var idxMovCred = -1;

    for (var c = 0; c < headers.length; c++) {
      if (headers[c].indexOf("movim. devedor") !== -1) idxMovDev = c;
      if (headers[c].indexOf("movim. credor") !== -1) idxMovCred = c;
    }

    if (idxConta === -1 || idxUg === -1) {
      throw new Error("Colunas obrigatórias ('Conta Contábil', 'UG Executora') não foram encontradas.");
    }

    var agrupado = {};

    for (var i = headerRowIndex + 1; i < data.length; i++) {
      var row = data[i];
      var conta = String(row[idxConta]).trim();
      
      // Regra 1: Descarta registros da conta 123119905 ou conta vazia
      if (conta === "123119905" || !conta) continue;

      var obsTexto = idxObs !== -1 ? String(row[idxObs]).trim() : "";
      
      // Regra 2: Extrai estritamente o número de 10 dígitos da transferência
      var numDoc = extrairNumDocTransferencia(obsTexto);
      
      // Regra 3: Se não for um documento de transferência legítimo, descarta o registro
      if (!numDoc) continue;

      var ug = String(row[idxUg]).trim();
      var devedor = parseFloat(row[idxMovDev]) || 0;
      var credor = parseFloat(row[idxMovCred]) || 0;

      // Extração do número da NL (12 últimos dígitos da coluna F)
      var valColF = (idxDoc !== -1 && row[idxDoc] !== undefined) ? String(row[idxDoc]).trim() : "";
      var nl = valColF.length >= 12 ? valColF.slice(-12) : valColF;

      // Estrutura para armazenar arrays de saídas e entradas
      if (!agrupado[numDoc]) {
        agrupado[numDoc] = {
          numDocumento: numDoc,
          saidas: [],
          entradas: []
        };
      }

      if (devedor > 0) {
        agrupado[numDoc].entradas.push({ conta: conta, ug: ug, valor: devedor, nl: nl });
      } else if (credor > 0) {
        agrupado[numDoc].saidas.push({ conta: conta, ug: ug, valor: credor, nl: nl });
      }
    }

    var resultado = Object.keys(agrupado).map(function(key) {
      return agrupado[key];
    });

    return { sucesso: true, dados: resultado };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Extrai unicamente o número do documento de transferência de 10 dígitos contido na observação.
 */
function extrairNumDocTransferencia(obs) {
  if (!obs) return null;
  
  // Procura o padrão "TRANSFERENCIA ... DE Nº: 2026400479" (exatos 10 dígitos)
  var match = obs.match(/TRANSFER[ÊE]NCIA\s+(?:INTERNA|EXTERNA)?.*?DE\s+N[ºo\.:\s]*(\d{10})/i);
  if (match && match[1]) {
    return match[1];
  }

  // Segunda verificação para o padrão "DE Nº: XXXXXXXXXX" com 10 dígitos
  var matchGen = obs.match(/DE\s+N[ºo\.:\s]*(\d{10})/i);
  if (matchGen && matchGen[1]) {
    return matchGen[1];
  }

  return null;
}

/**
 * Salva/Grava os dados de transferências processados na aba DOCUMENTO_MOVIMENTO.
 * @param {Array} dados Array de objetos contendo os dados das transferências
 */
function salvarDadosTransferencia(dados) {
  try {
    if (!dados || dados.length === 0) {
      throw new Error("Nenhum dado para salvar.");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DOCUMENTO_MOVIMENTO");

    if (!sheet) {
      throw new Error("A aba 'DOCUMENTO_MOVIMENTO' não foi encontrada na planilha.");
    }

    // Obtém todos os dados da planilha para rápida busca em memória
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      throw new Error("A aba 'DOCUMENTO_MOVIMENTO' não possui dados suficientes.");
    }

    // Leitura da Coluna A (Nº Documento)
    var colA = sheet.getRange(1, 1, lastRow, 1).getValues();
    
    // Leitura do bloco da Coluna N (14) até S (19) -> 6 colunas
    // Col N (14): NL Saída | Col O (15): Conta Saída | Col P (16): Valor Saída
    // Col Q (17): NL Entrada | Col R (18): Conta Entrada | Col S (19): Valor Entrada
    var rangeBloco = sheet.getRange(1, 14, lastRow, 6);
    var blocoValores = rangeBloco.getValues();

    // Mapeia o Nº do Documento de Transferência para o número da linha (índice 0-based)
    var mapaLinhas = {};
    for (var r = 0; r < colA.length; r++) {
      var numDocSheet = String(colA[r][0]).trim();
      if (numDocSheet) {
        mapaLinhas[numDocSheet] = r;
      }
    }

    var atualizados = 0;

    // Processa cada transferência da lista
    dados.forEach(function(item) {
      var numDoc = String(item.numDocumento).trim();
      
      if (mapaLinhas.hasOwnProperty(numDoc)) {
        var rowIndex = mapaLinhas[numDoc];

        var saidas = item.saidas || [];
        var entradas = item.entradas || [];

        // Monta as strings concatenadas com '|'
        var nlSaida = saidas.map(function(s) { return s.nl || ''; }).join('|');
        var contaSaida = saidas.map(function(s) { return s.conta || ''; }).join('|');
        var valorSaida = saidas.map(function(s) { return s.valor !== undefined ? s.valor : ''; }).join('|');

        var nlEntrada = entradas.map(function(e) { return e.nl || ''; }).join('|');
        var contaEntrada = entradas.map(function(e) { return e.conta || ''; }).join('|');
        var valorEntrada = entradas.map(function(e) { return e.valor !== undefined ? e.valor : ''; }).join('|');

        // Atualiza a matriz de dados nas colunas N, O, P, Q, R, S
        blocoValores[rowIndex][0] = nlSaida;      // Coluna N
        blocoValores[rowIndex][1] = contaSaida;   // Coluna O
        blocoValores[rowIndex][2] = valorSaida;   // Coluna P
        blocoValores[rowIndex][3] = nlEntrada;    // Coluna Q
        blocoValores[rowIndex][4] = contaEntrada;  // Coluna R
        blocoValores[rowIndex][5] = valorEntrada;  // Coluna S

        atualizados++;
      }
    });

    // Grava todas as alterações de uma só vez na planilha para alta performance
    rangeBloco.setValues(blocoValores);

    return { sucesso: true, quantidade: atualizados };

  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}