// 🌟 SUBSTITUA ESTA FUNÇÃO INTEGRALMENTE NO SEU MOVIMENTO.GS
// 🌟 COPIE E SUBSTITUA ESTA FUNÇÃO INTEGRALMENTE NO SEU MOVIMENTO.GS
function buscarTombosLote(texto) {
  const cache = CacheService.getScriptCache();
  const aba = SpreadsheetApp.getActive().getSheetByName("TOMBOS");
  const dados = aba.getDataRange().getValues();

  const linhas = texto.split("\n")
    .map(t => t.trim())
    .filter(t => t);

  const resultado = [];
  const jaAdicionados = new Set();

  linhas.forEach(tombo => {
    if (jaAdicionados.has(tombo)) return;

    let cacheKey = "tombo_" + tombo;
    cache.remove(cacheKey); // Limpa resíduos de cache antigo

    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];

      const siads = String(linha[0]).trim(); // Coluna A - Índice 0
      const sipac = String(linha[4]).trim(); // Coluna E - Índice 4

      if (sipac === tombo || siads === tombo) {
        const ugOrigemDoBem = String(linha[5]).trim();   // Coluna F - Índice 5
        const uorgOrigemDoBem = String(linha[3]).trim(); // Coluna D - Índice 3
        const descricaoDoBem = String(linha[2]).trim();  // Coluna C - Índice 2
        
        // 🌟 NOVO CAMINHO: Lê diretamente da Coluna G (Índice 6) que você criou
        const uorgPatrimonioDireta = String(linha[6] || "").trim();

        const item = {
          siads: siads,
          sipac: sipac,
          descricao: descricaoDoBem,
          uorg: uorgOrigemDoBem, 
          ug: ugOrigemDoBem,     
          // Se a coluna G estiver vazia na linha, avisa o usuário
          uorgPatrimonio: uorgPatrimonioDireta ? "✅ " + uorgPatrimonioDireta : "❌ Coluna G Vazia"
        };

        jaAdicionados.add(tombo);
        resultado.push(item);
        break;
      }
    }
  });

  return resultado;
}

// 🌟 FUNÇÃO DE BUSCA DA UORG DE PATRIMÔNIO TRATADA CONTRA ESPAÇOS EM BRANCO
function buscarUorgPatrimonio(ugOrigem) {
  if (!ugOrigem || String(ugOrigem).trim() === "" || String(ugOrigem).trim() === "-") return "";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("UORGS DE PATRIMONIO");
  if (!sheet) return "";
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return "";

  // Lê Coluna A (UG) e Coluna B (UORG Patrimônio)
  const dados = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  
  // Normalização para aceitar "00023" e número 23 como iguais
  const ugBuscaNorm = isNaN(ugOrigem) ? String(ugOrigem).trim() : Number(ugOrigem);
  
  for (let i = 0; i < dados.length; i++) {
    const ugPlanilha = dados[i][0];
    const ugPlanilhaNorm = isNaN(ugPlanilha) ? String(ugPlanilha).trim() : Number(ugPlanilha);
    
    if (ugBuscaNorm === ugPlanilhaNorm) {
      return String(dados[i][1]).trim(); // Retorna o valor da Coluna B (UORG PATRIMONIO)
    }
  }
  
  return ""; 
}

// 🌟 COPIE E SUBSTITUA NO SEU MOVIMENTO.GS (Garanta que só existe esta versão)
// 🌟 ATUALIZE TAMBÉM ESTA FUNÇÃO CASO ELA EXISTA NO SEU ARQUIVO
function carregarIndice() {
  if (INDEX) return INDEX;

  const aba = SpreadsheetApp.getActive().getSheetByName("TOMBOS");
  const dados = aba.getDataRange().getValues();
  INDEX = {};

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    
    const siads = String(linha[0]).trim(); 
    const uorg  = String(linha[3]).trim(); 
    const sipac = String(linha[4]).trim(); 
    const ug    = String(linha[5]).trim(); 
    
    // 🌟 NOVO CAMINHO: Coluna G (Índice 6)
    const uorgPatrimonioDireta = String(linha[6] || "").trim();

    const obj = {
      sipac: sipac,
      siads: siads,
      descricao: linha[2], 
      uorg: uorg,
      ug: ug,
      uorgPatrimonio: uorgPatrimonioDireta ? "✅ " + uorgPatrimonioDireta : "Não cadastrada"
    };

    if (sipac) INDEX[sipac] = obj;
    if (siads) INDEX[siads] = obj;
  }

  return INDEX;
}

function obterProximoSequencialMovimento() {

  const aba = SpreadsheetApp.getActive().getSheetByName("MOVIMENTO");
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 2) return 1;

  const valores = aba.getRange(2, 1, ultimaLinha - 1, 1).getValues();

  let maior = 0;

  valores.forEach(linha => {

    let v = linha[0];

    if (!v) return;

    // limpa espaços e força string
    v = String(v).trim();

    // pega tudo antes do ponto
    const base = v.split(".")[0];

    const numero = parseInt(base, 10);

    if (!isNaN(numero) && numero > maior) {
      maior = numero;
    }

  });

  return maior + 1;
}


function salvarMovimento(dados) {

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const aba = SpreadsheetApp.getActive().getSheetByName("MOVIMENTO");

    const sequencialBase = obterProximoSequencialMovimento();

    const linhas = [];
    const gruposComSequencial = [];

    dados.grupos.forEach((g, index) => {

      const sequencial = `${sequencialBase}.${index + 1}`;

      linhas.push([
        sequencial,
        g.tipoMov,
        dados.dataEnvioDBM,
        dados.processo,
        g.quantidade,
        g.ugOrigem,
        g.ugDestino,
        g.uorgOrigem,
        g.uorgDestino,
        "",
        "",
        "",
        "",
        "",
        "",
        "NÃO",
        ""
      ]);

      gruposComSequencial.push({
        sequencial,
        ...g
      });

    });

    if (linhas.length > 0) {
      aba.getRange(aba.getLastRow() + 1, 1, linhas.length, linhas[0].length)
         .setValues(linhas);
    }

    // 🔥 GERA PDF DENTRO DO LOCK (garante consistência)
    const pdfBase64 = gerarPdfMovimento({
     ...dados,
      bens: dados.bens || []
    }, gruposComSequencial);

    return {
      status: "ok",
      sequencialBase,
      total: linhas.length,
      pdf: pdfBase64
    };

  } finally {
    lock.releaseLock();
  }
}
function gerarPdfMovimento(dados, grupos) {

  const html = montarHtmlMovimento(dados, grupos);

  const blob = Utilities.newBlob(html, "text/html")
    .getAs("application/pdf")
    .setName("Movimentacao_" + dados.processo + ".pdf");

  return Utilities.base64Encode(blob.getBytes());
}

function montarHtmlMovimento(dados, grupos) {

  let linhasGrupos = "";
  let linhasBens = "";

  // 🔹 GRUPOS
  grupos.forEach(g => {
    linhasGrupos += `
      <tr>
        <td>${g.grupo}</td>
        <td>${g.sequencial}</td>
        <td>${g.ugOrigem}</td>
        <td>${g.uorgOrigem}</td>
        <td>${g.quantidade}</td>
        <td>${g.ugDestino}</td>
        <td>${g.uorgDestino}</td>
        <td>${g.tipoMov}</td>
      </tr>
    `;
  });

  // 🔹 BENS
  (dados.bens || []).forEach(b => {
    linhasBens += `
      <tr>
        <td>${b.sipac}</td>
        <td>${b.siads}</td>
        <td>${b.descricao}</td>
        <td>${b.ug}</td>
        <td>${b.uorg}</td>
      </tr>
    `;
  });

  return `
    <html>
    <head>
      <style>
        body {
          font-family: Arial;
          font-size: 12px;
        }

        h2 {
          text-align: center;
        }

        table {
          width: 100%;
          border-collapse: collapse;
          margin-top: 15px;
        }

        th {
          background: #1F4E78;
          color: white;
          padding: 6px;
        }

        td {
          border: 1px solid #ccc;
          padding: 5px;
          text-align: center;
        }

        .bloco {
          margin-top: 10px;
        }
      </style>
    </head>

    <body>

      <h2>Resumo de Movimentação de Bens</h2>

      <div class="bloco">
        <b>Processo:</b> ${dados.processo} <br>
        <b>Data Envio DBM:</b> ${dados.dataEnvioDBM} <br>
        <b>Tipo de Movimentação:</b> ${dados.tipoMovimentacao || ""}
      </div>

      <h3>Grupos de Movimentação</h3>

      <table>
        <tr>
          <th>GRUPO</th>
          <th>SEQ. DOC</th>
          <th>UG ORIGEM</th>
          <th>UORG ORIGEM</th>
          <th>QTD</th>
          <th>UG DESTINO</th>
          <th>UORG DESTINO</th>
          <th>TIPO</th>
        </tr>

        ${linhasGrupos}
      </table>

      <h3>Bens Movimentados</h3>

      <table>
        <tr>
          <th>SIPAC</th>
          <th>SIADS</th>
          <th>Descrição</th>
          <th>UG</th>
          <th>UORG</th>
        </tr>

        ${linhasBens}
      </table>

    </body>
    </html>
  `;
}

function obterDadosLinha(index) {

  const aba = SpreadsheetApp.getActive().getSheetByName("TOMBOS");
  const linha = aba.getRange(index + 1, 1, 1, 10).getValues()[0];

  return {
    sipac: String(linha[7]).trim(),
    siads: String(linha[8]).trim(),
    descricao: linha[2],
    ug: linha[9],
    uorg: linha[5]
  };
}

let INDEX = null;



// 🌟 ADICIONA A MESMA SEGURANÇA NA SEGUNDA FUNÇÃO DE BUSCA DO SERVIDOR
function buscarTombosMovimento(listaTombos) {
  const index = carregarIndice();

  return listaTombos.map(tombo => {
    const item = index[tombo];
    if (item) return item;

    return {
      sipac: tombo,
      siads: "NÃO ENCONTRADO",
      descricao: "Tombo não localizado na base de dados",
      ug: "-",
      uorg: "-",
      uorgPatrimonio: "Não cadastrada"
    };
  });
}

function salvarProcessoRecolhimento(dados, arquivo) {
  const pasta = DriveApp.getFolderById("1thPL3SiIw_6Ia_bN4jHS_kPBP5gYo0I2");
  let linkArquivo = "";

  if (arquivo) {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(arquivo.base64),
      arquivo.mimeType,
      arquivo.nome
    );
    const subpasta = pasta.createFolder(dados.processo);
    const file = subpasta.createFile(blob);
    linkArquivo = file.getUrl();
  }

  const aba = SpreadsheetApp
    .openById("1qrUFrqUnYvi51FoUfmAYdV2h6lOMh3fHM3-tjDz3Zyc")
    .getSheetByName("RECOLHIMENTO");

  const usuario = dados.usuario || Session.getActiveUser().getEmail() || "NÃO IDENTIFICADO";
  const dataProcesso = dados.dataProcesso ? new Date(dados.dataProcesso) : "";

  const linha = [[
    String(dados.processo).trim(), // A (1)
    dados.ug,                      // B (2)
    dados.departamento,            // C (3)
    dataProcesso,                  // D (4)
    dados.quantidade,              // E (5)
    usuario,                       // F (6)
    new Date(),                    // G (7)
    "",                            // H (8)
    "",                            // I (9)
    "",                            // J (10)
    "",                            // K (11)
    "PENDENTE DE AGENDAMENTO",     // L (12)
    "",                            // M (13) -> Vazio inicialmente (reservado para a data)
    linkArquivo                    // N (14) -> 🔥 O LINK AGORA ENTRA AQUI COMPATÍVEL COM A DATA EM M
  ]];

  aba.getRange(aba.getLastRow() + 1, 1, 1, linha[0].length).setValues(linha);
  return { status: "ok", linkArquivo };
}

function buscarProcessoRecolhimento(processo) {

  const procBusca = normalizarProcesso(processo);

  const sheet = SpreadsheetApp
    .openById("1qrUFrqUnYvi51FoUfmAYdV2h6lOMh3fHM3-tjDz3Zyc")
    .getSheetByName("RECOLHIMENTO");

  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {

    const procLinha = normalizarProcesso(dados[i][0]);

    if (procLinha === procBusca) {
      return {
        ug: dados[i][1],
        departamento: dados[i][2],
        qtd: dados[i][4] // ⚠️ ajuste aqui (explico abaixo)
      };
    }
  }

  return null;
}

function normalizarProcesso(valor) {
  return String(valor)
    .replace(/\u00A0/g, "")   // remove espaço invisível (NBSP)
    .replace(/\s/g, "")       // remove TODOS espaços
    .replace(/[^\d]/g, "")    // mantém só números
    .trim();
}

function getListaDestinos() {

  const aba = SpreadsheetApp.getActive()
    .getSheetByName("LISTAS");

  const dados = aba.getRange(2, 4, aba.getLastRow() - 1, 2).getValues();
  // Coluna D (destino) e E (UORG)

  return dados
    .filter(l => l[0]) // remove vazios
    .map(l => ({
      destino: l[0],
      uorg: l[1]
    }));
}

function testeLista() {
  Logger.log(getListaDestinos());
}

function salvarRecolhimento(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaMov = ss.getSheetByName("DOCUMENTO_MOVIMENTO");
  const abaRec = ss.getSheetByName("RECOLHIMENTO");
  
  // ================= 1. NOVA LÓGICA DE SEQUENCIAL (10 DÍGITOS) =================
  const anoAtual = new Date().getFullYear().toString(); // Captura o ano corrente (Ex: 2026)
  const digitoFixo = "4";                               // O 5º dígito definido
  const prefixoBusca = anoAtual + digitoFixo;          // Padrão inicial de busca (Ex: "20264")
  
  let proximoIncremento = 1;
  const ultimaLinhaAba = abaMov.getLastRow();
  
  if (ultimaLinhaAba > 1) {
    // Busca todos os sequenciais existentes na Coluna A
    const valoresSeq = abaMov.getRange("A2:A" + ultimaLinhaAba).getValues();
    let maiorIncrementoAno = 0;
    
    valoresSeq.forEach(r => {
      const seqStr = String(r[0]).trim();
      // Valida se o sequencial pertence ao lote do ano atual e tem o tamanho correto de 10 dígitos
      if (seqStr.startsWith(prefixoBusca) && seqStr.length === 10) {
        const incrementoStr = seqStr.slice(5); // Isola os 5 últimos dígitos
        const incrementoNum = parseInt(incrementoStr, 10);
        if (!isNaN(incrementoNum) && incrementoNum > maiorIncrementoAno) {
          maiorIncrementoAno = incrementoNum;
        }
      }
    });
    
    proximoIncremento = maiorIncrementoAno + 1;
  }
  
  const linhasParaGravar = [];

  // ================= 2. MONTAGEM DAS LINHAS COM O NOVO PADRÃO =================
  dados.grupos.forEach(g => {
    // Formata o incremento para ocupar 5 dígitos com zeros à esquerda (Ex: 00001, 00002...)
    const incrementoFormatado = String(proximoIncremento).padStart(5, '0');
    
    // Monta o Sequencial Final de 10 dígitos
    const sequencialFinal = Number(prefixoBusca + incrementoFormatado);
    
    // Guarda o sequencial dentro do objeto do grupo para o PDF e relatórios usarem depois
    g.sequencialGerado = sequencialFinal; 

    linhasParaGravar.push([
      sequencialFinal,        // A - Sequencial (Ex: 2026400001)
      "RECOLHIMENTO",         // B - Tipo Doc
      g.tipoMov,              // C - Tipo Movimento
      dados.processo,         // D - Nº DO PROCESSO
      g.quantidade,           // E - Qtd
      g.tombosStr,            // F - Tombos
      g.ugOrigem,             // G - UG Origem
      g.ugDestino,            // H - UG Destino
      g.uorgOrigem,           // I - UORG Origem
      g.uorgDestino,          // J - UORG Destino
      g.uorgPatrimonio || ""  // K - UORG PATRIMÔNIO
    ]);
    
    // Incrementa para o próximo grupo do mesmo loop (se houver mais de um grupo no mesmo envio)
    proximoIncremento++;
  });

  // 3. Gravação física na planilha
  if (linhasParaGravar.length > 0) {
    abaMov.getRange(abaMov.getLastRow() + 1, 1, linhasParaGravar.length, 11).setValues(linhasParaGravar);
  }

  // 4. Atualização das datas na aba RECOLHIMENTO
  try {
    const dadosRec = abaRec.getDataRange().getValues();
    dados.grupos.forEach(g => {
      g.bens.forEach(bem => {
        const tomboProc = String(bem.sipac || bem.siads).trim();
        for (let i = 1; i < dadosRec.length; i++) {
          if (String(dadosRec[i][0]).trim() === tomboProc) {
            abaRec.getRange(i + 1, 3).setValue(dados.dataColeta); // Grava data na Coluna C
            break;
          }
        }
      });
    });
  } catch (e) {
    console.error("Erro ao atualizar aba RECOLHIMENTO: " + e.message);
  }

  // 5. Geração do PDF chamando a função atualizada
  const pdfBase64 = gerarPdfRecolhimento(dados);

  return {
    status: "ok",
    sequencialBase: linhasParaGravar[linhasParaGravar.length - 1][0], // Retorna o último sequencial gerado real
    pdf: pdfBase64
  };
}

function gerarPdfRecolhimento(dados) {
  // Estilização básica para um PDF profissional
  let html = `
    <html>
      <head>
        <style>
          body { font-family: sans-serif; margin: 20px; color: #333; }
          .header { border-bottom: 2px solid #1f4e78; margin-bottom: 20px; padding-bottom: 10px; }
          .header h2 { color: #1f4e78; margin: 0; }
          .info { margin-bottom: 20px; font-size: 11px; }
          table { width: 100%; border-collapse: collapse; font-size: 9px; }
          th, td { border: 1px solid #000; padding: 5px; text-align: left; }
          th { background-color: #f2f2f2; font-weight: bold; }
          .seq-col { background-color: #fafafa; font-weight: bold; width: 60px; }
          .tombos-cell { word-break: break-all; max-width: 250px; }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>RESUMO DE RECOLHIMENTO DE BENS</h2>
        </div>
        
        <div class="info">
          <b>Processo SIPAC:</b> ${dados.processo}<br>
          <b>Data da Coleta:</b> ${dados.dataColeta}<br>
          <b>Gerado em:</b> ${Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm")}
        </div>
        
        <table>
          <thead>
            <tr>
              <th>SEQUENCIAL</th>
              <th>UG ORIGEM</th>
              <th>UORG ORIGEM</th>
              <th>UORG PATRIMÔNIO</th>
              <th>UORG DESTINO</th>
              <th>QTD</th>
              <th>TOMBOS</th>
            </tr>
          </thead>
          <tbody>
  `;

  dados.grupos.forEach(g => {
    html += `
      <tr>
        <td class="seq-col">${g.sequencialGerado}</td>
        <td>${g.ugOrigem}</td>
        <td>${g.uorgOrigem}</td>
        <td>${g.uorgPatrimonio || "N/A"}</td>
        <td>${g.uorgDestino}</td>
        <td>${g.quantidade}</td>
        <td class="tombos-cell">${g.tombosStr}</td>
      </tr>
    `;
  });

  html += `
          </tbody>
        </table>
        <div style="margin-top: 30px; font-size: 10px; text-align: center;">
          <p>_________________________________________________________</p>
          <p>Assinatura do Responsável</p>
        </div>
      </body>
    </html>
  `;

  // Processo de conversão para PDF real (Blob)
  const htmlOutput = HtmlService.createHtmlOutput(html);
  const blob = htmlOutput.getAs('application/pdf').setName("Recolhimento_" + dados.processo + ".pdf");

  return Utilities.base64Encode(blob.getBytes());
}


function testarPermissaoDrive() {
  DriveApp.getFolderById("109OGnX_GkAjin3ICwtSO_KQh-yZoGhK4");
  Logger.log("Permissão OK!");
}


function salvarAbaTransferencia(dados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("TRANSFERENCIA");
    
    if (!aba) throw new Error("Aba TRANSFERENCIA não encontrada!");

    let urlArquivoDrive = ""; // Armazenará o link do arquivo caso seja feito o upload

    // 🔥 LOGICA DE CRIAÇÃO E ARMAZENAMENTO DO ARQUIVO NO GOOGLE DRIVE
    if (dados.arquivo && dados.arquivo.base64) {
      const idPastaPai = "109OGnX_GkAjin3ICwtSO_KQh-yZoGhK4"; // ID retirado da sua URL
      const pastaPai = DriveApp.getFolderById(idPastaPai);
      
      // Sanitiza o número do processo para usar como nome de pasta
      const nomeSubpasta = "Processo_" + String(dados.processo).replace(/[^0-9]/g, "");
      
      // Verifica se a subpasta do processo já existe para evitar duplicados
      let pastaProcesso;
      const pastasExistentes = pastaPai.getFoldersByName(nomeSubpasta);
      if (pastasExistentes.hasNext()) {
        pastaProcesso = pastasExistentes.next();
      } else {
        pastaProcesso = pastaPai.createFolder(nomeSubpasta);
      }
      
      // Converte o código base64 de volta em um arquivo físico
      const blob = Utilities.newBlob(
        Utilities.base64Decode(dados.arquivo.base64), 
        dados.arquivo.mimeType, 
        dados.arquivo.nome
      );
      
      // Salva o arquivo dentro da pasta do processo
      const arquivoCriado = pastaProcesso.createFile(blob);
      urlArquivoDrive = arquivoCriado.getUrl(); // Captura o link definitivo do arquivo
    }

    // Regra de Status: SIM -> Pendente de agendamento | NÃO -> Finalizado
    const statusDefinido = (dados.carga === "SIM") ? "Pendente de agendamento" : "Finalizado";
    
    // 🔥 NOVA REGRA AUTOMÁTICA PARA UG DE ORIGEM (COLUNA J)
    let ugOrigemFinal = dados.ugOrigem || "";
    if (dados.tipo) {
      const tipoTexto = String(dados.tipo).trim().toUpperCase();
      
      // Valida se o tipo é Distribuição de Bens (com ou sem acento) ou Reaproveitamento
      if (
        tipoTexto.includes("DISTRIB") || 
        tipoTexto.includes("REAPROVEITAMENTO")
      ) {
        ugOrigemFinal = "153409"; // Força a UG de Origem padrão
      }
    }
    
    // Mapeamento exato das colunas da aba TRANSFERENCIA incluindo o link do arquivo na Coluna L
    aba.appendRow([
      dados.tipo,             // A
      dados.carga,            // B
      dados.processo,         // C
      dados.dataEnvio,        // D
      dados.ugDestino,        // E
      dados.uorgDestino,      // F
      getUsuarioLogado(),     // G
      new Date(),             // H
      statusDefinido,         // I
      ugOrigemFinal,          // J - 🔥 Agora utiliza a variável tratada pela regra acima
      dados.uorgOrigem || "", // K
      urlArquivoDrive,        // L - Link da Solicitação de Movimentação no Drive
      
      // 🌟 COLUNAS GRAVADAS EM LOTE (M ao R)
      dados.servidorRetirada || "",  // M
      dados.telefoneRetirada || "",  // N
      dados.emailRetirada || "",     // O
      dados.servidorEntrega || "",   // P
      dados.telefoneEntrega || "",   // Q
      dados.emailEntrega || ""       // R
    ]);
    
    return "OK";

  } catch (e) {
    throw new Error(e.message);
  }
}

// Mantemos a função de sequencial para uso posterior na Aba 2
function obterNovoSequencialTransferencia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("DOCUMENTO_MOVIMENTO");
  const ultimaLinha = aba.getLastRow();
  
  // 1. Definição do padrão dinâmico com o ano atual e o dígito fixo "2"
  const anoAtual = new Date().getFullYear().toString(); // Captura o ano corrente (Ex: 2026)
  const digitoFixo = "2";                               // O 5º dígito definido para Transferência
  const prefixoBusca = anoAtual + digitoFixo;          // Padrão inicial de busca (Ex: "20262")
  
  let proximoIncremento = 1;
  
  // 2. Varredura inteligente na Coluna A se a planilha já contiver dados
  if (ultimaLinha > 1) {
    const valores = aba.getRange("A2:A" + ultimaLinha).getValues();
    let maiorIncrementoAno = 0;
    
    valores.forEach(r => {
      const seqStr = String(r[0]).trim();
      // Valida se o sequencial pertence ao lote do ano atual e tem o tamanho correto de 10 dígitos
      if (seqStr.startsWith(prefixoBusca) && seqStr.length === 10) {
        const incrementoStr = seqStr.slice(5); // Isola os 5 últimos dígitos
        const incrementoNum = parseInt(incrementoStr, 10);
        if (!isNaN(incrementoNum) && incrementoNum > maiorIncrementoAno) {
          maiorIncrementoAno = incrementoNum;
        }
      }
    });
    
    proximoIncremento = maiorIncrementoAno + 1;
  }
  
  // 3. Formata o incremento com zeros à esquerda (5 dígitos) e monta o Sequencial Final
  const incrementoFormatado = String(proximoIncremento).padStart(5, '0');
  const sequencialFinal = Number(prefixoBusca + incrementoFormatado);
  
  return sequencialFinal;
}

/**
 * Localiza um processo na aba TRANSFERENCIA para validar a movimentação
 */
function buscarProcessoTransferencia(processoProcurado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("TRANSFERENCIA");
    if (!aba) return null;

    const dados = aba.getDataRange().getValues();
    
    // Procura o processo na Coluna C (índice 2)
    // r[2] é a coluna do Processo
    const linha = dados.find(r => String(r[2]).trim() === String(processoProcurado).trim());

    if (!linha) return null;

    return {
      tipo: linha[0],        // Coluna A
      ugDestino: linha[4],   // Coluna E
      uorgDestino: linha[5]  // Coluna F
    };
    
  } catch (e) {
    throw new Error("Erro ao buscar processo: " + e.message);
  }
}

function salvarTombosTransferencia(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("DOCUMENTO_MOVIMENTO");
    if (!aba) throw new Error("Aba DOCUMENTO_MOVIMENTO não encontrada!");
    const dataAtual = new Date();
    const usuario = getUsuarioLogado();
    const resultadosSequenciais = [];
    const gruposParaPdf = [];
    
    // 1. Agrupar os itens novamente no servidor para garantir integridade
    const grupos = {};
    payload.itens.forEach(item => {
      const chave = `${item.ug}_${item.uorg}`;
      if (!grupos[chave]) {
        grupos[chave] = { 
          ugOrigem: item.ug, 
          uorgOrigem: item.uorg, 
          tombos: [], 
          qtd: 0,
          // 🌟 CAPTURA A UORG PATRIMÔNIO LIMPA QUE FOI EXIBIDA NO RESUMO DA TELA
          uorgPatrimonio: item.uorgPatrimonio ? String(item.uorgPatrimonio).replace("✅", "").trim() : ""
        };
      }
      grupos[chave].tombos.push(item.siads);
      grupos[chave].qtd++;
    });

    let contadorGrupo = 1;

    // 2. Salvar cada grupo como uma linha (sequencial) diferente
    for (let chave in grupos) {
      const g = grupos[chave];
      const novoSequencial = obterNovoSequencialTransferencia();

      // Regra de classificação SIADS
      const tipoSiads = (String(g.ugOrigem).trim() === String(payload.ugDestino).trim()) ?
        "TROCA DE RESPONSABILIDADE" : "TRANSFERENCIA ENTRE UG'S";

      // 🌟 SE A UORG PATRIMÔNIO NÃO VEIO TRATADA DO BANCO, EXECUTA A BUSCA AUTOMÁTICA VIA FUNÇÃO INTERNA
      let uorgPatrimonioFinal = g.uorgPatrimonio;
      if (!uorgPatrimonioFinal || uorgPatrimonioFinal.includes("❌") || uorgPatrimonioFinal === "") {
        uorgPatrimonioFinal = buscarUorgPatrimonio(g.ugOrigem);
      }

      // GRAVAÇÃO FÍSICA CORRIGIDA (Alinhando perfeitamente as colunas A até K)
      aba.appendRow([
        novoSequencial,       // A: SEQUENCIAL DOCUMENTO
        payload.tipo,         // B: TIPO DE MOVIMENTAÇÃO
        tipoSiads,            // C: TIPO DE MOVIMENTAÇÃO SIADS
        payload.processo,     // D: PROCESSO
        g.qtd,                // E: QUANTIDADE DE BENS
        g.tombos.join("; "),  // F: TOMBAMENTOS
        g.ugOrigem,           // G: UG DE ORIGEM
        payload.ugDestino,    // H: UG DE DESTINO
        g.uorgOrigem,         // I: UORG DE ORIGEM
        payload.uorgDestino,  // J: UORG DE DESTINO
        uorgPatrimonioFinal,   // K: 🌟 CORRIGIDO! Salva a UORG de Patrimônio real em vez da origem normal
        usuario,              // L: 🌟 NOVO: USUÁRIO RESPONSÁVEL (E-mail)
        dataAtual             // M: 🌟 NOVO: DATA E HORA DO SALVAMENTO
      ]);

      // Monta os objetos de retorno para o fluxo seguir normalmente...
      resultadosSequenciais.push(novoSequencial);
      
      gruposParaPdf.push({
        grupo: "Grupo " + contadorGrupo,
        sequencial: novoSequencial,
        tipoMov: tipoSiads,
        quantidade: g.qtd,
        ugOrigem: g.ugOrigem,
        uorgOrigem: g.uorgOrigem,
        uorgPatrimonio: uorgPatrimonioFinal, // 🌟 Alinha também no PDF gerado
        ugDestino: payload.ugDestino,
        uorgDestino: payload.uorgDestino
      });
      
      contadorGrupo++;
    }

    const sequenciaisTexto = resultadosSequenciais.join(", ");

    // 3. GERA O RELATÓRIO EM PDF
    let pdfBase64 = null;
    try {
      // Formata os dados gerais do cabeçalho
      const dadosParaPdf = {
        processo: payload.processo,
        dataEnvioDBM: Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        tipoMovimentacao: payload.tipo,
        bens: payload.itens // Mapeia a lista de tombos
      };

      // 🔥 CHAMA A SUA FUNÇÃO PASSANDO OS DOIS ARGUMENTOS REQUERIDOS
      pdfBase64 = gerarPdfMovimento(dadosParaPdf, gruposParaPdf);
      
    } catch (errPdf) {
      console.error("Erro interno ao gerar o PDF:", errPdf);
    }

    // 4. Retorna a resposta estruturada para o Frontend
    return {
      sucesso: true,
      sequencialGerado: sequenciaisTexto,
      pdf: pdfBase64
    };

  } catch (e) {
    throw new Error("Erro no servidor: " + e.message);
  }
}

/**
 * Exemplo ilustrativo de suporte (ajuste conforme o nome/código real do seu gerador de PDF):
 */
function gerarPdfBase64Transferencia(dados) {
  // Aqui dentro roda a lógica padrão do seu sistema que cria o HTML do relatório,
  // transforma em Blob e depois faz o: Utilities.base64Encode(pdfBlob.getBytes());
  // Certifique-se de que possui essa função operando corretamente no backend.
  
  if (typeof gerarPdfMovimento === "function") {
     return gerarPdfMovimento(dados); 
  }
  return null;
}


// Função para RECOLHIMENTO
function buscarDadosParaMac(sequencial) {
  return processarBuscaGenerica(sequencial);
}

// Função para TRANSFERÊNCIA
function buscarDadosParaMacMov(sequencial) {
  return processarBuscaGenerica(sequencial);
}

// Função auxiliar interna para evitar repetição
function processarBuscaGenerica(sequencial) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName("DOCUMENTO_MOVIMENTO");
    const dados = aba.getDataRange().getValues();
    const linha = dados.find(r => String(r[0]) === String(sequencial));

    if (!linha) return null;

    return {
      sequencial: linha[0],       // Coluna A
      tipoMovimentacao: linha[2],  // Coluna C (Tipo de movimentação SIADS)
      processo: linha[3],         // Coluna D
      tombos: linha[5],           // Coluna F
      ugOrigem: linha[6],           // Coluna G
      ugDestino: linha[7],        // 🌟 CORRIGIDO: Puxa a UG de Destino da Coluna H (Índice 7)
      uorgOrigem: linha[8],       // Coluna I
      uorgDestino: linha[9],      // Coluna J
      uorgPatrimonio: linha[10]   // 🔥 NOVO: Coluna K (UORG de Patrimônio na Origem)
    };
  } catch (e) {
    throw new Error("Erro na busca: " + e.message);
  }
}

/**
 * Busca todas as movimentações do tipo "TRANSFERENCIA ENTRE UG'S" filtradas por mês e ano.
 * Alinhado perfeitamente com a estrutura real da aba DOCUMENTO_MOVIMENTO
 * @param {Object} filtros - Contém { mes: "01" a "12", ano: "2026" }
 */
/**
 * Busca todas as movimentações do tipo "TRANSFERENCIA ENTRE UG'S" filtradas por mês e ano.
 */
function buscarMovimentacoesParaNL(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaMov = ss.getSheetByName("DOCUMENTO_MOVIMENTO");
    if (!abaMov) return { sucesso: false, message: "Aba 'DOCUMENTO_MOVIMENTO' não encontrada." };

    const dados = abaMov.getDataRange().getValues();
    const contas = obterContasListaB(); 

    if (dados.length <= 1) return { sucesso: true, movimentacoes: [], contas: contas };

    const mesFiltro = parseInt(filtros.mes);
    const anoFiltro = parseInt(filtros.ano);
    const result = [];

    for (let i = 1; i < dados.length; i++) {
      const r = dados[i];
      
      // Coluna C (Índice 2): TIPO DE MOVIMENTAÇÃO SIADS
      let tipoOriginal = String(r[2] || "").trim();
      let tipoSiads = tipoOriginal.toUpperCase();
      tipoSiads = tipoSiads.replace(/[ÁÀÂÃ]/g, "A").replace(/[ÉÈÊ]/g, "E").replace(/[ÍÌÎ]/g, "I").replace(/[ÓÒÔÕ]/g, "O").replace(/[ÚÙÛ]/g, "U");

      if (tipoSiads.indexOf("TRANSFERENCIA ENTRE UG") !== -1) {
        
        // Coluna M (Índice 12): DATA E HORA
        let celulaData = r[12]; 
        let dataMov = null;
        let dataFormatada = "-";

        if (celulaData instanceof Date) {
          dataMov = celulaData;
          // Formata a data para DD/MM/AAAA para exibir bonito na tabela
          dataFormatada = Utilities.formatDate(dataMov, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else if (celulaData) {
          dataMov = new Date(celulaData);
          if (!isNaN(dataMov.getTime())) {
            dataFormatada = Utilities.formatDate(dataMov, Session.getScriptTimeZone(), "dd/MM/yyyy");
          } else {
            dataFormatada = String(celulaData).split(" ")[0]; // Pega apenas a data se for texto
          }
        }

        if (dataMov && !isNaN(dataMov.getTime())) {
          const anoMov = dataMov.getFullYear();
          const mesMov = dataMov.getMonth() + 1;

          if (anoMov !== anoFiltro || mesMov !== mesFiltro) {
            continue; 
          }
        } else {
          continue; 
        }

        result.push({
          linhaReal: i + 1,
          sequencial: String(r[0] || "").trim(),   // Coluna A
          tipo: tipoOriginal,                      // Coluna C 🌟 NOVA
          processo: String(r[3] || "").trim(),     // Coluna D
          ugOrigem: String(r[6] || "").trim(),     // Coluna G
          ugDestino: String(r[7] || "").trim(),    // Coluna H
          data: dataFormatada,                     // Coluna M 🌟 NOVA (Formatada)
          
          nlSaida: String(r[13] || "").trim(),     // Coluna N
          contaSaida: String(r[14] || "").trim(),  // Coluna O
          valorSaida: String(r[15] || "").trim(),  // Coluna P
          nlEntrada: String(r[16] || "").trim(),   // Coluna Q
          contaEntrada: String(r[17] || "").trim(),// Coluna R
          valorEntrada: String(r[18] || "").trim() // Coluna S
        });
      }
    }

    return { sucesso: true, movimentacoes: result, contas: contas };

  } catch (e) {
    return { sucesso: false, message: "Erro no processamento do Servidor: " + e.message };
  }
}

/**
 * Busca a relação de contas contábeis na aba LISTAS, coluna B
 */
function obterContasListaB() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaListas = ss.getSheetByName("LISTAS");
    if (!abaListas) return [];
    
    const uLinha = abaListas.getLastRow();
    if (uLinha < 2) return [];
    
    const valores = abaListas.getRange(2, 2, uLinha - 1, 1).getValues();
    return valores.map(v => String(v[0]).trim()).filter(v => v !== "");
  } catch(e) {
    return [];
  }
}

/**
 * Grava em lote os dados de NL de Movimentação na aba DOCUMENTO_MOVIMENTO 
 * e atualiza automaticamente a UORG e a UG dos bens correspondentes na aba TOMBOS
 */
function salvarNLMovimentacao(loteDados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaMov = ss.getSheetByName("DOCUMENTO_MOVIMENTO");
    const abaTombos = ss.getSheetByName("TOMBOS");
    
    if (!abaMov) return { sucesso: false, message: "Aba 'DOCUMENTO_MOVIMENTO' não encontrada." };

    // Carrega os dados da aba TOMBOS para a memória (Performance rápida)
    let dadosTombos = abaTombos ? abaTombos.getDataRange().getValues() : null;

    loteDados.forEach(item => {
      const vLinha = parseInt(item.linhaReal, 10);
      
      // 1. Grava os dados de NL nas colunas N a S (Índices físicos 14 a 19)
      abaMov.getRange(vLinha, 14).setValue(String(item.nlSaida).toUpperCase());   // Coluna N
      abaMov.getRange(vLinha, 15).setValue(String(item.contaSaida));              // Coluna O
      abaMov.getRange(vLinha, 16).setValue(String(item.valorSaida));              // Coluna P
      abaMov.getRange(vLinha, 17).setValue(String(item.nlRecebimento).toUpperCase()); // Coluna Q
      abaMov.getRange(vLinha, 18).setValue(String(item.contaRecebimento));        // Coluna R
      abaMov.getRange(vLinha, 19).setValue(String(item.valorRecebimento));        // Coluna S

      // 2. ATUALIZAÇÃO DA UORG E UG NA ABA TOMBOS (Se a aba existir)
      if (abaTombos && dadosTombos) {
        // Captura os dados armazenados na linha atual da aba DOCUMENTO_MOVIMENTO (Colunas A até L)
        const valoresLinha = abaMov.getRange(vLinha, 1, 1, 12).getValues()[0];
        
        const txtTombamentos = String(valoresLinha[5]).trim(); // Coluna F (TOMBAMENTOS) - Índice 5
        const ugDestino = String(valoresLinha[7]).trim();      // Coluna H (UG DE DESTINO) - Índice 7 🔥 NOVO
        const uorgDestino = String(valoresLinha[9]).trim();    // Coluna J (UORG DE DESTINO) - Índice 9

        if (txtTombamentos && (uorgDestino || ugDestino)) {
          // Quebra o texto da coluna de tombos por vírgula, espaço ou quebra de linha
          const listaTombosAlvo = txtTombamentos.split(/[\s,\n]+/)
            .map(t => t.trim())
            .filter(t => t.length > 0);

          if (listaTombosAlvo.length > 0) {
            const conjuntoTombos = new Set(listaTombosAlvo);

            // Varre a matriz da aba TOMBOS procurando os bens correspondentes
            for (let k = 1; k < dadosTombos.length; k++) {
              const siadsTombos = String(dadosTombos[k][0]).trim(); // Coluna A (tombo_SIADS) - Índice 0
              const sipacTombos = String(dadosTombos[k][4]).trim(); // Coluna E (tombo_SIPAC) - Índice 4

              // Se o tombo da planilha coincidir com a lista da movimentação, altera os dados de destino
              if (conjuntoTombos.has(siadsTombos) || conjuntoTombos.has(sipacTombos)) {
                
                // Grava a nova UORG de destino na Coluna D (Índice 3) da aba TOMBOS
                if (uorgDestino) {
                  abaTombos.getRange(k + 1, 4).setValue(uorgDestino);
                }
                
                // Grava a nova UG de destino na Coluna F (Índice 5) da aba TOMBOS 🔥 NOVO
                if (ugDestino) {
                  abaTombos.getRange(k + 1, 6).setValue(ugDestino);
                }
              }
            }
          }
        }
      }
    });

    return { sucesso: true };
  } catch(e) {
    return { sucesso: false, message: "Erro ao salvar NL e atualizar dados do Bem: " + e.message };
  } finally {
    lock.releaseLock();
  }
}


// 🔥 ATALHOS DE SEGURANÇA PARA EVITAR ERROS DE REFERENCE_ERROR / CACHE DO NAVEGADOR
function obterMovimentacoesNL(filtros) { return listarMovimentacoesNL(filtros); }
function obterMovimentacoesParaNL(filtros) { return listarMovimentacoesNL(filtros); }
