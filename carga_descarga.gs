function extrairApenasNumeros(str) {
  return String(str || "").replace(/\D/g, "");
}

function normalizarUorgTexto(str) {
  const s = String(str || "").trim();
  if (!s) return "";
  const digitos = s.replace(/\D/g, "").replace(/^0+/, "");
  return digitos || s.toLowerCase();
}

function buscarSequencialDocMov(dadosDocMov, numProc, uorgDestino, ocorrenciaTransf) {
  if (!dadosDocMov || dadosDocMov.length === 0) return "";
  
  const procLimpo = extrairApenasNumeros(numProc);
  if (!procLimpo) return "";
  
  const uorgLimpa = normalizarUorgTexto(uorgDestino);
  const matches = [];

  for (let d = 1; d < dadosDocMov.length; d++) {
    const seqDocMov = String(dadosDocMov[d][0] || "").trim();      // Coluna A (Sequencial)
    const procDocMov = String(dadosDocMov[d][3] || "").trim();     // Coluna D (Processo)
    const uorgDestDocMov = String(dadosDocMov[d][9] || "").trim(); // Coluna J (UORG Destino)

    if (extrairApenasNumeros(procDocMov) === procLimpo) {
      matches.push({
        sequencial: seqDocMov,
        uorgNorm: normalizarUorgTexto(uorgDestDocMov)
      });
    }
  }

  if (matches.length === 0) return "";
  if (matches.length === 1) return matches[0].sequencial;

  // Tenta vinculação exata/parcial por UORG
  if (uorgLimpa) {
    const uorgMatches = matches.filter(m => 
      m.uorgNorm && (m.uorgNorm === uorgLimpa || m.uorgNorm.includes(uorgLimpa) || uorgLimpa.includes(m.uorgNorm))
    );
    if (uorgMatches.length === 1) return uorgMatches[0].sequencial;
  }

  // Se houver mais de um sequencial para o mesmo processo, pega pela ordem de ocorrência
  const idx = Math.min(ocorrenciaTransf, matches.length - 1);
  return matches[idx].sequencial;
}

function obterPainelCargaDescarga() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaUg = ss.getSheetByName("UG");
  const abaRec = ss.getSheetByName("RECOLHIMENTO");
  const abaTransf = ss.getSheetByName("TRANSFERENCIA");
  const abaDocMov = ss.getSheetByName("DOCUMENTO_MOVIMENTO");
  
  if (!abaUg || !abaRec || !abaTransf) throw new Error("Abas necessárias não localizadas.");

  const dadosUg = abaUg.getDataRange().getValues();
  const dadosRec = abaRec.getDataRange().getValues();
  const dadosTransf = abaTransf.getDataRange().getValues();
  const dadosDocMov = abaDocMov ? abaDocMov.getDataRange().getValues() : [];
  
  let painel = {};
  
  for (let i = 1; i < dadosUg.length; i++) {
    const codigoUg = String(dadosUg[i][0]).replace(/\s+/g, '').trim(); 
    const descricaoUg = String(dadosUg[i][1]).trim();
    const siglaUg = String(dadosUg[i][2]).trim();
    if (!codigoUg) continue;
    
    painel[codigoUg] = {
      codigo: codigoUg,
      nome: siglaUg ? siglaUg + " - " + descricaoUg : descricaoUg,
      recolhimentos: [],
      transferencias: []
    };
  }
  
  // RECOLHIMENTO
  for (let r = 1; r < dadosRec.length; r++) {
    const linha = dadosRec[r];
    const status = String(linha[11] || "").replace(/\s+/g, ' ').trim().toUpperCase(); 
    const ugProcesso = String(linha[1] || "").replace(/\s+/g, '').trim(); 
    
    if (status === "PENDENTE DE AGENDAMENTO" && painel[ugProcesso]) {
      painel[ugProcesso].recolhimentos.push({
        linhaReal: r + 1,
        processo: String(linha[0] || "").trim(),     
        sequencial: "-",                             
        ug: ugProcesso,                              
        departamento: String(linha[2] || "").trim(), 
        dataProcesso: formatarDataGsheet(linha[3]),   
        quantBens: String(linha[4] || "").trim()     
      });
    }
  }
  
  // TRANSFERENCIA
  let contadorProcTransf = {};

  for (let t = 1; t < dadosTransf.length; t++) {
    const linha = dadosTransf[t];
    const cargaDescarga = String(linha[1] || "").replace(/\s+/g, '').trim().toUpperCase();
    const status = String(linha[8] || "").replace(/\s+/g, ' ').trim().toUpperCase();
    const ugOrigem = String(linha[9] || "").replace(/\s+/g, '').trim(); 
    const numProc = String(linha[2] || "").trim();
    const uorgDestino = String(linha[5] || "").trim();
    
    if (cargaDescarga === "SIM" && status === "PENDENTE DE AGENDAMENTO" && painel[ugOrigem]) {
      const procLimpo = extrairApenasNumeros(numProc);
      const idxOcorrencia = contadorProcTransf[procLimpo] || 0;
      contadorProcTransf[procLimpo] = idxOcorrencia + 1;

      let seqDoc = String(linha[6] || "").trim();
      
      if (!seqDoc || seqDoc.includes("@") || seqDoc === "-") {
        seqDoc = buscarSequencialDocMov(dadosDocMov, numProc, uorgDestino, idxOcorrencia);
      }

      painel[ugOrigem].transferencias.push({
        linhaReal: t + 1,
        sequencial: seqDoc || "-",
        processo: numProc,
        tipoTransferencia: String(linha[0] || "").trim(),
        ugOrigem: ugOrigem,
        uorgOrigem: String(linha[10] || "").trim(),
        dataEnvioDbm: formatarDataGsheet(linha[3])
      });
    }
  }
  
  return Object.values(painel).filter(u => u.recolhimentos.length > 0 || u.transferencias.length > 0);
}

// Função auxiliar para buscar a quantidade de bens em DOCUMENTO_MOVIMENTO
function obterQtdBensDocMov(dadosDocMov, sequencial, numProcesso) {
  if (!dadosDocMov || dadosDocMov.length <= 1) return "-";

  const seqLimpo = extrairApenasNumeros(sequencial);
  const procLimpo = extrairApenasNumeros(numProcesso);

  // 1ª Tentativa: Busca pelo Sequencial na Coluna A (índice 0)
  if (seqLimpo) {
    for (let i = 1; i < dadosDocMov.length; i++) {
      const seqLinha = extrairApenasNumeros(dadosDocMov[i][0]);
      if (seqLinha && seqLinha === seqLimpo) {
        const valColE = String(dadosDocMov[i][4] || "").trim(); // Coluna E
        if (valColE && valColE !== "0") return valColE;

        // Se Coluna E estiver vazia, conta os tombos da Coluna F (índice 5)
        const valColF = String(dadosDocMov[i][5] || "").trim();
        if (valColF) {
          const qtdContada = valColF.split(";").map(t => t.trim()).filter(Boolean).length;
          if (qtdContada > 0) return String(qtdContada);
        }
      }
    }
  }

  // 2ª Tentativa: Busca pelo Número do Processo na Coluna D (índice 3)
  if (procLimpo) {
    for (let i = 1; i < dadosDocMov.length; i++) {
      const procLinha = extrairApenasNumeros(dadosDocMov[i][3]);
      if (procLinha && procLinha === procLimpo) {
        const valColE = String(dadosDocMov[i][4] || "").trim();
        if (valColE && valColE !== "0") return valColE;

        const valColF = String(dadosDocMov[i][5] || "").trim();
        if (valColF) {
          const qtdContada = valColF.split(";").map(t => t.trim()).filter(Boolean).length;
          if (qtdContada > 0) return String(qtdContada);
        }
      }
    }
  }

  return "-";
}

function obterAgendaLogistica(dataFiltroString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaUg = ss.getSheetByName("UG");
  const abaRec = ss.getSheetByName("RECOLHIMENTO");
  const abaTransf = ss.getSheetByName("TRANSFERENCIA");
  const abaDocMov = ss.getSheetByName("DOCUMENTO_MOVIMENTO");
  
  if (!abaUg || !abaRec || !abaTransf) throw new Error("Abas necessárias não localizadas.");

  const dadosUg = abaUg.getDataRange().getValues();
  const dadosRec = abaRec.getDataRange().getValues();
  const dadosTransf = abaTransf.getDataRange().getValues();
  const dadosDocMov = abaDocMov ? abaDocMov.getDataRange().getValues() : [];

  let dicioUg = {};
  for (let i = 1; i < dadosUg.length; i++) {
    const cod = String(dadosUg[i][0]).replace(/\s+/g, '').trim();
    const desc = String(dadosUg[i][1]).trim();
    const sigla = String(dadosUg[i][2]).trim();
    if (cod) dicioUg[cod] = sigla ? `${sigla} - ${desc}` : desc;
  }

  let dataFiltroFormatada = "";
  if (dataFiltroString && dataFiltroString.includes("-")) {
    const partes = dataFiltroString.split("-");
    dataFiltroFormatada = partes[2] + "/" + partes[1] + "/" + partes[0];
  }

  let agenda = [];
  const converterParaTextoDdmmyyyy = (valorRaw) => {
    if (!valorRaw) return "";
    if (valorRaw instanceof Date) return Utilities.formatDate(valorRaw, Session.getScriptTimeZone(), "dd/MM/yyyy");
    return String(valorRaw).trim();
  };

  // RECOLHIMENTO
  for (let r = 1; r < dadosRec.length; r++) {
    const status = String(dadosRec[r][11] || "").trim().toUpperCase();
    const dataRaw = dadosRec[r][12]; 
    const codUg = String(dadosRec[r][1] || "").replace(/\s+/g, '').trim();

    if (status === "AGENDADO" && dataRaw) {
      const dataTexto = converterParaTextoDdmmyyyy(dataRaw);
      if (dataFiltroFormatada && dataTexto !== dataFiltroFormatada) continue;

      const qtdRec = String(dadosRec[r][4] || "").trim() || "-";

      agenda.push({
        linhaReal: r + 1,
        tipo: "RECOLHIMENTO",
        data: dataTexto,
        sequencial: "-",
        processo: String(dadosRec[r][0] || "").trim(),
        ugDescricao: dicioUg[codUg] || `UG ${codUg} (Não cadastrada)`,
        quantBens: qtdRec,
        qtdBens: qtdRec
      });
    }
  }

  // TRANSFERENCIA / DISTRIBUIÇÃO
  let contadorProcTransf = {};

  for (let t = 1; t < dadosTransf.length; t++) {
    const status = String(dadosTransf[t][8] || "").trim().toUpperCase();
    const dataRaw = dadosTransf[t][11]; 
    const codUgOrigem = String(dadosTransf[t][9] || "").replace(/\s+/g, '').trim();
    const tipoMovimentacao = String(dadosTransf[t][0] || "").trim().toUpperCase();
    const numProc = String(dadosTransf[t][2] || "").trim();
    const uorgDestino = String(dadosTransf[t][5] || "").trim();

    const procLimpo = extrairApenasNumeros(numProc);
    const idxOcorrencia = contadorProcTransf[procLimpo] || 0;
    contadorProcTransf[procLimpo] = idxOcorrencia + 1;

    if (status === "AGENDADO" && dataRaw) {
      const dataTexto = converterParaTextoDdmmyyyy(dataRaw);
      if (dataFiltroFormatada && dataTexto !== dataFiltroFormatada) continue;

      const ehDist = tipoMovimentacao.includes("DISTRIB") || tipoMovimentacao.includes("BENS");
      const tipoDefinido = ehDist ? "DISTRIBUICAO_BENS" : "TRANSFERENCIA";

      let seqDoc = String(dadosTransf[t][6] || "").trim();
      if (!seqDoc || seqDoc.includes("@") || seqDoc === "-") {
        seqDoc = buscarSequencialDocMov(dadosDocMov, numProc, uorgDestino, idxOcorrencia);
      }

      // Busca a quantidade na Coluna E ou faz a contagem da Coluna F em DOCUMENTO_MOVIMENTO
      const qtdCalculada = obterQtdBensDocMov(dadosDocMov, seqDoc, numProc);

      agenda.push({
        linhaReal: t + 1,
        tipo: tipoDefinido,
        data: dataTexto,
        sequencial: seqDoc || "-",
        processo: numProc,
        ugDescricao: dicioUg[codUgOrigem] || `UG ${codUgOrigem} (Não cadastrada)`,
        quantBens: qtdCalculada,
        qtdBens: qtdCalculada
      });
    }
  }

  return agenda.sort((a, b) => {
    let pA = a.data.split("/"), pB = b.data.split("/");
    return new Date(pA[2], pA[1]-1, pA[0]) - new Date(pB[2], pB[1]-1, pB[0]);
  });
}

function obterLinkOuGerarGuia(linhaReal, tipo, processo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vLinha = parseInt(linhaReal, 10);
    
    if (tipo === "RECOLHIMENTO") {
      const aba = ss.getSheetByName("RECOLHIMENTO");
      const url = aba.getRange(vLinha, 14).getValue();
      if (!url) return { sucesso: false, mensagem: "Nenhum documento foi anexado a este processo de Recolhimento." };
      return { sucesso: true, tipo: "LINK", url: url };
    } 

    const abaTransf = ss.getSheetByName("TRANSFERENCIA");
    const dadosLinha = abaTransf.getRange(vLinha, 1, 1, 12).getValues()[0];
    
    const tipoTransf      = String(dadosLinha[0] || "");
    const numProcesso     = String(dadosLinha[2] || "").trim();
    const dataEnvioRaw    = dadosLinha[3];
    const ugDestinoCod    = String(dadosLinha[4] || "").replace(/\s+/g, '').trim();
    
    let uorgDestinoCod  = String(dadosLinha[5] || "").replace(/\s+/g, '').trim();
    if (uorgDestinoCod) uorgDestinoCod = uorgDestinoCod.padStart(6, '0');
    
    let sequencialDoc   = String(dadosLinha[6] || "").trim();
    const ugOrigemCod     = String(dadosLinha[9] || "").replace(/\s+/g, '').trim();
    
    let uorgOrigemCod   = String(dadosLinha[10] || "").replace(/\s+/g, '').trim();
    if (uorgOrigemCod) uorgOrigemCod = uorgOrigemCod.padStart(6, '0');
    
    const urlAnexo        = String(dadosLinha[11] || ""); 
    
    if (urlAnexo && urlAnexo.startsWith("http")) {
      return { sucesso: true, tipo: "LINK", url: urlAnexo };
    }
    
    const formatarData = (d) => (d instanceof Date) ? Utilities.formatDate(d, "GMT-3", "dd/MM/yyyy") : String(d);

    const abaDocMov = ss.getSheetByName("DOCUMENTO_MOVIMENTO");
    const dadosDocMov = abaDocMov ? abaDocMov.getDataRange().getValues() : [];

    // Calcula índice de ocorrência da linha na aba TRANSFERENCIA
    const procLimpo = extrairApenasNumeros(numProcesso);
    const todosDadosTransf = abaTransf.getDataRange().getValues();
    let idxOcorrencia = 0;
    for (let t = 1; t < vLinha - 1; t++) {
      if (extrairApenasNumeros(todosDadosTransf[t][2]) === procLimpo) {
        idxOcorrencia++;
      }
    }

    // Se o Sequencial for e-mail ou inválido, realiza a busca exata em DOCUMENTO_MOVIMENTO
    if (!sequencialDoc || sequencialDoc.includes("@") || sequencialDoc === "-") {
      sequencialDoc = buscarSequencialDocMov(dadosDocMov, numProcesso, uorgDestinoCod, idxOcorrencia);
    }

    // Extrai tombos estritamente do Sequencial localizado
    let listaTombamentos = [];
    if (abaDocMov && sequencialDoc && sequencialDoc !== "-") {
      for (let i = 1; i < dadosDocMov.length; i++) {
        const lDoc = dadosDocMov[i];
        const seqDoc = String(lDoc[0] || "").trim();   // Coluna A = SEQUENCIAL
        const tomboRaw = String(lDoc[5] || "").trim(); // Coluna F = TOMBOS

        if (seqDoc === sequencialDoc && tomboRaw) {
          const tombosSplitted = tomboRaw.split(";");
          tombosSplitted.forEach(t => {
            const tomboLimpo = t.trim();
            if (tomboLimpo !== "" && !listaTombamentos.includes(tomboLimpo)) {
              listaTombamentos.push(tomboLimpo);
            }
          });
        }
      }
    }

    // MAPEAR INFORMAÇÕES DOS BENS (Aba TOMBOS)
    const abaTombos = ss.getSheetByName("TOMBOS");
    let dicioBens = {};
    if (abaTombos) {
      const dadosTombos = abaTombos.getDataRange().getValues();
      for (let i = 1; i < dadosTombos.length; i++) {
        const tomboSiads = String(dadosTombos[i][0] || "").trim();
        const nomeMaterial = String(dadosTombos[i][2] || "").trim();
        const tomboSipac = String(dadosTombos[i][4] || "").trim();
        
        if (tomboSiads) {
          dicioBens[tomboSiads] = {
            siads: tomboSiads,
            sipac: tomboSipac || "-",
            descricao: nomeMaterial || "Material sem descrição cadastrada"
          };
        }
      }
    }

    // MONTAGEM DA TABELA DE BENS
    let linhasTabelaBensHtml = "";
    let contadorBens = 0;

    listaTombamentos.forEach(tRef => {
      contadorBens++;
      const bem = dicioBens[tRef] || {
        siads: tRef,
        sipac: "-",
        descricao: "Material sem descrição cadastrada"
      };

      linhasTabelaBensHtml += `
        <tr>
          <td style="text-align: center;">${contadorBens}</td>
          <td style="font-weight: bold; color: #2c3e50;">${bem.siads}</td>
          <td>${bem.sipac}</td>
          <td>${bem.descricao}</td>
        </tr>`;
    });

    if (linhasTabelaBensHtml === "") {
      linhasTabelaBensHtml = `<tr><td colspan="4" style="text-align: center; color: #7f8c8d; font-style: italic; padding: 15px;">Nenhum bem patrimonial localizado para o sequencial ${sequencialDoc || "-"}.</td></tr>`;
    }

    // Nomes de UG e UORG
    const abaUg = ss.getSheetByName("UG");
    let dicioUg = {};
    if (abaUg) {
      const dadosUg = abaUg.getDataRange().getValues();
      for (let i = 1; i < dadosUg.length; i++) {
        const cod = String(dadosUg[i][0]).replace(/\s+/g, '').trim();
        const desc = String(dadosUg[i][1]).trim();
        if (cod) dicioUg[cod] = desc;
      }
    }
    const nomeUgOrigem  = dicioUg[ugOrigemCod]  ? `${ugOrigemCod} - ${dicioUg[ugOrigemCod]}`   : `UG ${ugOrigemCod}`;
    const nomeUgDestino = dicioUg[ugDestinoCod] ? `${ugDestinoCod} - ${dicioUg[ugDestinoCod]}` : `UG ${ugDestinoCod}`;

    const abaOrganograma = ss.getSheetByName("ORGANOGRAMA");
    let dicioUorg = {};
    if (abaOrganograma) {
      const dadosOrg = abaOrganograma.getDataRange().getValues();
      dadosOrg.forEach((lOrg, idx) => {
        if (idx === 0) return;
        let cSipac = String(lOrg[0] || "").replace(/\s+/g, '').trim().padStart(6, '0');
        let dUorg = String(lOrg[1] || "").trim();
        let cSiads = String(lOrg[2] || "").replace(/\s+/g, '').trim().padStart(6, '0');
        if (dUorg) {
          if (cSipac) dicioUorg[cSipac] = dUorg;
          if (cSiads) dicioUorg[cSiads] = dUorg;
        }
      });
    }
    const nomeUorgOrigem  = dicioUorg[uorgOrigemCod]  ? `${uorgOrigemCod} - ${dicioUorg[uorgOrigemCod]}`   : (uorgOrigemCod || "Não especificada");
    const nomeUorgDestino = dicioUorg[uorgDestinoCod] ? `${uorgDestinoCod} - ${dicioUorg[uorgDestinoCod]}` : (uorgDestinoCod || "Não especificada");

    let htmlGuia = `
      <html>
      <head>
        <style>
          body { font-family: 'Arial', sans-serif; margin: 25px; color: #2c3e50; line-height: 1.4; }
          .topo-guia { text-align: center; border-bottom: 2px solid #2c3e50; padding-bottom: 12px; margin-bottom: 25px; }
          .topo-guia h1 { margin: 0; font-size: 14px; color: #1a252f; text-transform: uppercase; font-weight: bold; letter-spacing: 0.5px; line-height: 1.5; }
          .titulo-documento { text-align: center; font-size: 13px; font-weight: bold; background: #f2f4f4; padding: 6px; margin-bottom: 20px; text-transform: uppercase; border: 1px solid #bdc3c7; }
          .secao-dados { margin-bottom: 15px; }
          .secao-titulo { font-size: 11px; font-weight: bold; background: #34495e; color: white; padding: 4px 8px; margin-bottom: 8px; text-transform: uppercase; }
          .grid-dados { width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 11px; }
          .grid-dados td { padding: 6px 8px; border: 1px solid #d5dbdb; vertical-align: top; }
          .label { font-weight: bold; color: #34495e; display: block; font-size: 9px; text-transform: uppercase; margin-bottom: 2px; }
          .valor { font-size: 11px; color: #000; }
          .tabela-bens { width: 100%; border-collapse: collapse; font-size: 11px; margin-top: 5px; }
          .tabela-bens th { background: #eaeded; color: #2c3e50; font-weight: bold; border: 1px solid #bdc3c7; padding: 6px; text-transform: uppercase; font-size: 9px; }
          .tabela-bens td { padding: 6px; border: 1px solid #d5dbdb; vertical-align: middle; }
          .bloco-assinaturas { margin-top: 55px; width: 100%; }
          .bloco-assinaturas td { width: 50%; text-align: center; font-size: 10px; border: none; }
          .linha-assinatura { width: 75%; margin: 0 auto 4px auto; border-top: 1px solid #7f8c8d; }
          .rodape { position: fixed; bottom: 0; left: 0; width: 100%; text-align: center; font-size: 8px; color: #95a5a6; border-top: 1px solid #bdc3c7; padding-top: 4px; }
        </style>
      </head>
      <body>
        <div class="topo-guia">
          <h1>PRÓ-REITORIA DE ADMINISTRAÇÃO<br>DIRETORIA DE GESTÃO PATRIMONIAL E DOCUMENTAL</h1>
        </div>
        <div class="titulo-documento">GUIA DE MOVIMENTAÇÃO</div>
        
        <div class="secao-dados">
          <div class="secao-titulo">Controle do Processo Administrativo</div>
          <table class="grid-dados">
            <tr>
              <td style="width: 30%;"><span class="label">Sequencial Doc.</span><span class="valor"><strong>${sequencialDoc || "-"}</strong></span></td>
              <td style="width: 40%;"><span class="label">Nº do Processo</span><span class="valor"><strong>${numProcesso}</strong></span></td>
              <td style="width: 30%;"><span class="label">Tipo de Movimentação</span><span class="valor">${tipoTransf}</span></td>
            </tr>
            <tr>
              <td colspan="3"><span class="label">Data do Envio</span><span class="valor">${formatarData(dataEnvioRaw)}</span></td>
            </tr>
          </table>
        </div>

        <div class="secao-dados">
          <div class="secao-titulo">Origem / Remetente</div>
          <table class="grid-dados">
            <tr>
              <td style="width: 40%;"><span class="label">Unidade Gestora (UG)</span><span class="valor">${nomeUgOrigem}</span></td>
              <td style="width: 60%;"><span class="label">Unidade Administrativa (UORG)</span><span class="valor">${nomeUorgOrigem}</span></td>
            </tr>
          </table>
        </div>

        <div class="secao-dados">
          <div class="secao-titulo">Destino / Recebedor</div>
          <table class="grid-dados">
            <tr>
              <td style="width: 40%;"><span class="label">Unidade Gestora (UG)</span><span class="valor">${nomeUgDestino}</span></td>
              <td style="width: 60%;"><span class="label">Unidade Administrativa (UORG)</span><span class="valor">${nomeUorgDestino}</span></td>
            </tr>
          </table>
        </div>

        <div class="secao-dados">
          <div class="secao-titulo">Relação Descritiva de Bens Patrimoniais</div>
          <table class="tabela-bens">
            <thead>
              <tr>
                <th style="width: 6%; text-align: center;">Item</th>
                <th style="width: 22%; text-align: left;">Tombo SIADS</th>
                <th style="width: 22%; text-align: left;">Tombo SIPAC</th>
                <th style="text-align: left;">Descrição do Bem Material</th>
              </tr>
            </thead>
            <tbody>
              ${linhasTabelaBensHtml}
            </tbody>
          </table>
        </div>

        <table class="bloco-assinaturas">
          <tr>
            <td><div class="linha-assinatura"></div>Responsável pela Unidade de Origem</td>
            <td><div class="linha-assinatura"></div>Responsável pela Unidade de Destino</td>
          </tr>
        </table>

        <div class="rodape">Guia emitida em ${Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss")}</div>
      </body>
      </html>`;

    const idNomePdf = sequencialDoc ? `Guia_${sequencialDoc}` : `Guia_${numProcesso.replace(/[^0-9]/g, "")}`;
    const blobHtml = Utilities.newBlob(htmlGuia, "text/html", idNomePdf + ".html");
    const blobPdf = blobHtml.getAs("application/pdf").setName(idNomePdf + ".pdf");
    
    return {
      sucesso: true,
      tipo: "PDF_GERADO",
      base64: Utilities.base64Encode(blobPdf.getBytes()),
      nomeArquivo: blobPdf.getName()
    };
  } catch(e) {
    return { sucesso: false, mensagem: "Erro ao compilar documento: " + e.message };
  }
}

function gravarAgendamentoPlanilha(linhaReal, tipo, dataString) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vLinha = parseInt(linhaReal, 10);
    const partes = dataString.split("-");
    const dataFormatada = partes[2] + "/" + partes[1] + "/" + partes[0];

    const tipoUpper = String(tipo).toUpperCase();

    if (tipoUpper === "RECOLHIMENTO") {
      const aba = ss.getSheetByName("RECOLHIMENTO");
      if (!aba) return { sucesso: false, mensagem: "Aba RECOLHIMENTO não encontrada." };
      aba.getRange(vLinha, 12).setValue("AGENDADO");
      aba.getRange(vLinha, 13).setValue(dataFormatada);
    } else {
      const aba = ss.getSheetByName("TRANSFERENCIA");
      if (!aba) return { sucesso: false, mensagem: "Aba TRANSFERENCIA não encontrada." };
      aba.getRange(vLinha, 9).setValue("AGENDADO");
      aba.getRange(vLinha, 12).setValue(dataFormatada);
    }

    return { sucesso: true, mensagem: "Agendamento logístico gravado com sucesso para " + dataFormatada + "!" };
  } catch(e) {
    return { sucesso: false, mensagem: "Erro ao gravar agendamento: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function formatarDataGsheet(dataRaw) {
  if (!dataRaw) return "-";
  if (dataRaw instanceof Date) {
    return Utilities.formatDate(dataRaw, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return String(dataRaw).trim();
}

function alterarDataAgendamento(linhaReal, tipo, novaDataString) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vLinha = parseInt(linhaReal, 10);
    const partes = novaDataString.split("-");
    const dataFormatada = partes[2] + "/" + partes[1] + "/" + partes[0];

    const nomeAba = tipo === "RECOLHIMENTO" ? "RECOLHIMENTO" : "TRANSFERENCIA";
    const colData = tipo === "RECOLHIMENTO" ? 13 : 12;
    
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) return { sucesso: false, mensagem: "Aba " + nomeAba + " não encontrada." };
    
    aba.getRange(vLinha, colData).setValue(dataFormatada);
    return { sucesso: true, mensagem: "Data alterada para " + dataFormatada + "!" };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  } finally { lock.releaseLock(); }
}

function excluirAgendamentoLogistico(linhaReal, tipo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vLinha = parseInt(linhaReal, 10);
    
    const nomeAba = tipo === "RECOLHIMENTO" ? "RECOLHIMENTO" : "TRANSFERENCIA";
    const colStatus = tipo === "RECOLHIMENTO" ? 12 : 9;
    const colData = tipo === "RECOLHIMENTO" ? 13 : 12;
    
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) return { sucesso: false, mensagem: "Aba " + nomeAba + " não encontrada." };
    
    aba.getRange(vLinha, colStatus).setValue("PENDENTE DE AGENDAMENTO");
    aba.getRange(vLinha, colData).setValue("");
    
    return { sucesso: true, mensagem: "Agendamento cancelado com sucesso." };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  } finally { lock.releaseLock(); }
}

function gerarPdfProgramacaoDia(dataFiltroString) {
  try {
    const servicos = obterAgendaLogistica(dataFiltroString);
    const partes = dataFiltroString.split("-");
    const dataFormatada = partes[2] + "/" + partes[1] + "/" + partes[0];
    
    if (servicos.length === 0) {
      return { sucesso: false, mensagem: "Não existem serviços agendados nesta data." };
    }

    let html = `
      <html>
      <head>
        <style>
          body { font-family: 'Helvetica', 'Arial', sans-serif; margin: 30px; color: #333; }
          .header { text-align: center; border-bottom: 2px solid #0056b3; padding-bottom: 15px; margin-bottom: 25px; }
          .header h1 { margin: 0; font-size: 20px; color: #0056b3; text-transform: uppercase; }
          .header h2 { margin: 5px 0 0 0; font-size: 14px; color: #555; font-weight: normal; }
          .meta-info { margin-bottom: 20px; font-size: 11px; text-align: right; color: #666; }
          table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 11px; }
          th { background-color: #f2f2f2; border: 1px solid #ddd; padding: 8px; text-align: left; font-weight: bold; color: #111; }
          td { border: 1px solid #ddd; padding: 8px; }
          .badge { padding: 3px 6px; border-radius: 3px; font-weight: bold; font-size: 9px; text-align: center; display: inline-block; }
          .badge-rec { background-color: #fff3cd; color: #856404; border: 1px solid #ffeeba; }
          .badge-transf { background-color: #d1ecf1; color: #0c5460; border: 1px solid #bee5eb; }
          .text-center { text-align: center; }
          .footer { position: fixed; bottom: 0; left: 0; width: 100%; text-align: center; font-size: 9px; color: #999; border-top: 1px solid #ddd; padding-top: 5px; }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>Relatório Institucional de Programação Logística</h1>
          <h2>Módulo Operacional de Carga e Descarga</h2>
        </div>
        
        <div class="meta-info">
          <strong>Data da Programação:</strong> ${dataFormatada} <br>
          <strong>Emitido em:</strong> ${Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss")}
        </div>

        <table>
          <thead>
            <tr>
              <th style="width: 15%;">Tipo Serviço</th>
              <th style="width: 15%;">Sequencial</th>
              <th style="width: 20%;">Nº do Processo</th>
              <th>Unidade Gestora (UG)</th>
              <th style="width: 12%;" class="text-center">Qtd Bens</th>
            </tr>
          </thead>
          <tbody>`;

    servicos.forEach(s => {
      const ehRec = s.tipo === "RECOLHIMENTO";
      const classeBadge = ehRec ? "badge-rec" : "badge-transf";
      
      html += `
        <tr>
          <td><span class="badge ${classeBadge}">${s.tipo}</span></td>
          <td><strong>${s.sequencial || '-'}</strong></td>
          <td><strong>${s.processo}</strong></td>
          <td>${s.ugDescricao}</td>
          <td class="text-center">${s.quantBens || '-'}</td>
        </tr>`;
    });

    html += `
          </tbody>
        </table>
        
        <div class="footer">
          Sistema Integrado de Gestão Patrimonial - Diretoria de Bens Móveis (DBM)
        </div>
      </body>
      </html>`;

    const blob = Utilities.newBlob(html, "text/html", "Relatorio_Programacao_" + dataFiltroString + ".html");
    const pdfBlob = blob.getAs("application/pdf").setName("Relatorio_Programacao_" + dataFormatada.replace(/\//g, "-") + ".pdf");
    
    return { sucesso: true, base64: Utilities.base64Encode(pdfBlob.getBytes()), nomeArquivo: pdfBlob.getName() };
  } catch(e) {
    return { sucesso: false, mensagem: "Erro ao gerar PDF: " + e.message };
  }
}