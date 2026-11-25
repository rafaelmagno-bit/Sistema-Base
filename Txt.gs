// =============================
// 📄 GERAÇÃO TXT ENTRADA (ESTRUTURA ATUALIZADA)
// =============================

function gerarTxtEntrada(termoBusca) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("ENTRADA");
  const dados = aba.getDataRange().getValues();
  
  const linhasFiltradas = dados.slice(1).filter(linha => {
    const valorProcesso = String(linha[5] || "");
    const valorSequencial = String(linha[0] || "");
    return valorProcesso === termoBusca || valorSequencial === termoBusca;
  });

  if (linhasFiltradas.length === 0) {
    throw new Error("Nenhum registro encontrado para o termo: " + termoBusca);
  }

  const primeiraLinha = linhasFiltradas[0];
  const dadosComuns = {
    processo: primeiraLinha[5],
    data: primeiraLinha[4],
    tipo: primeiraLinha[3],
    uorg: primeiraLinha[7]
  };

  const ENTER = "[enter]";
  const TAB = "[tab]";
  const BACKTAB = "[backtab]";
  
  // ============================================================
  // 1. COMANDO INICIAL: ACESSAR A TRANSAÇÃO
  // ============================================================
  let txt = BACKTAB + ">incentrada" + ENTER;

  // 2. Gerar blocos de comandos para cada item encontrado
  linhasFiltradas.forEach((linha, i) => {
    const cadmat = String(linha[23]).trim();
    const infoComplementar = String(linha[24] || "");
    const marcaOriginal = String(linha[25] || "");
    const valorUnitarioComFrete = linha[28]; 
    const qtdOriginal = linha[27]; 
    const qtdFormatada = String(qtdOriginal || 1).trim(); 
    
    // 🌟 CAPTURA DOS DADOS DO CORRESPONSÁVEL (Coluna P = índice 15, Coluna Q = índice 16)
    const nomeCorresponsavel = String(linha[15] || "").trim();
    const cpfCru = String(linha[16] || "").replace(/\D/g, ""); // Mantém apenas números
    const cpfFormatado = cpfCru.padStart(11, "0"); // Garante 11 dígitos com zeros à esquerda
    
    const valorLimpo = limparNumero(valorUnitarioComFrete);
    const dataFormatada = formatarDataTXT(dadosComuns.data);

    const descricaoBase = `${infoComplementar} - PROC ${dadosComuns.processo}`;
    const descricaoNormalizada = normalizarTexto(descricaoBase);
    const marcaNormalizada = normalizarTexto(marcaOriginal);

    const tabs = calcularTabs(descricaoNormalizada);
    if (tabs === "ERRO") {
      throw new Error(`Descrição excede 209 caracteres no item ${i + 1}`);
    }

    // Comandos de Inclusão por item
    txt += "X" + ENTER;
    txt += "2" + qtdFormatada + ENTER;
    txt += `${cadmat}${TAB}00001X` + ENTER;

    txt += descricaoNormalizada
      + tabs
      + `${dataFormatada}${valorLimpo}`
      + TAB + tipoEntrada(dadosComuns.tipo)
      + TAB + "1" 
      + TAB + TAB
      + marcaNormalizada
      + TAB + TAB
      + ENTER;

    // 🌟 ALTERAÇÃO SOLICITADA AQUI:
    // Reduzido para 4 TABs e concatenado o CPF + Nome + TAB
    const tabUorg = calcularTabUORG(dadosComuns.uorg);
    txt += TAB + TAB + TAB + TAB
      + cpfFormatado
      + nomeCorresponsavel
      + TAB
      + TAB
      + dadosComuns.uorg
      + tabUorg
      + "15233" 
      + ENTER;

    txt += "s" + ENTER;
    txt += "n" + ENTER;
  });

  // ============================================================
  // 3. COMANDO FINAL PARA GERAR RELATÓRIO
  // ============================================================
  const PF3 = "[pf3]";
  const uorg6Digitos = String(dadosComuns.uorg || "").trim().padStart(6, "0");

  let comandoRelatorio = PF3 + BACKTAB + ">relentrada" + ENTER;
  comandoRelatorio += TAB + TAB + uorg6Digitos;
  comandoRelatorio += TAB + TAB + TAB + TAB + TAB + TAB + TAB + TAB + TAB + TAB;
  comandoRelatorio += "06" + ENTER;

  txt += comandoRelatorio;

  return txt;
}

// --- FUNÇÕES AUXILIARES MANTIDAS E REVISADAS ---

function limparNumero(valor) {
  return Number(valor).toFixed(2).replace(".", "");
}

function formatarDataTXT(data) {
  if (!data) return "";
  const d = new Date(data);
  if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "ddMMyyyy");
}

function tipoEntrada(tipo) {
  if (!tipo) return "";

  // Normaliza o texto: remove acentos e deixa em maiúsculo
  // Ex: "Doação" vira "DOACAO"
  return normalizarTexto(tipo);
}

function normalizarTexto(txt) {
  if (!txt) return "";
  return txt
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/["'\[\]]/g, "")
    .replace(/[^\w\s-]/g, "") // Remove caracteres especiais mantendo espaços e hífens
    .toUpperCase();
}

function calcularTabs(descricao) {
  const len = descricao.length;
  if (len > 208) return "ERRO";
  if (len === 208) return "";
  if (len >= 133) return "[tab]";
  if (len >= 57) return "[tab][tab]";
  return "[tab][tab][tab]";
}

function calcularTabUORG(uorg) {
  const valor = String(uorg || "").trim();
  const len = valor.length;
  if (len > 6) return "ERRO";
  return len === 6 ? "" : "[tab]";
}

function gerarArquivoMAC(tr) {
  const conteudo = gerarTxtEntrada(tr);
  const xml = `<HAScript name="INCLUSAO_ENTRADA" description="" timeout="60000" pausetime="300" promptall="true" blockinput="false" author="Rafael" creationdate="${new Date().toLocaleString()}" supressclearevents="false" usevars="false" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true">
    <screen name="Tela1" entryscreen="true" exitscreen="true" transient="false">
        <description>
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="${conteudo}" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
    </screen>
</HAScript>`;

  return {
    nome: `ENTRADA_${tr.replace(/\//g, "_")}.mac`,
    mime: "text/plain",
    base64: Utilities.base64Encode(xml)
  };
}