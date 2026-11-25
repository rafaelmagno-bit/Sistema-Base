function salvarSaidaPainel(payload) {
  // Opcional: Descomente para travar o uso apenas para DBM ou ADMIN
  // const perfil = getPerfilUsuario();
  // if (perfil !== "ADMIN" && perfil !== "DBM") {
  //   throw new Error("Acesso negado. Perfil sem permissão para registrar saídas.");
  // }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName("SAIDAS");
  
  // Cria a aba se ela não existir por acidente na planilha
  if (!aba) {
    aba = ss.insertSheet("SAIDAS");
    aba.appendRow(["SEQUENCIAL DOCUMENTO", "DATA_REGISTRO", "USUARIO", "TIPO", "DATA DA SAÍDA", "PROCESSO", "QUANTIDADE DE TOMBAMENTOS", "TOMBOS"]);
  }

  const dados = aba.getDataRange().getValues();
  const anoAtual = new Date().getFullYear();
  
  // 1. GERAÇÃO DO SEQUENCIAL INTELIGENTE (10 DÍGITOS: ANO + 9 + SEQUENCIAL)
  let proximoSequencial = 1;
  const prefixoProcurado = String(anoAtual) + "9"; // Ex: "20269"

  for (let i = 1; i < dados.length; i++) {
    const seqExistente = String(dados[i][0]);
    if (seqExistente.startsWith(prefixoProcurado)) {
      const numeroFinal = parseInt(seqExistente.substring(5), 10);
      if (!isNaN(numeroFinal) && numeroFinal >= proximoSequencial) {
        proximoSequencial = numeroFinal + 1;
      }
    }
  }

  // Preenche com zeros à esquerda até o sequencial completar as 5 posições restantes (Ex: 00001)
  const seqFormatado = prefixoProcurado + String(proximoSequencial).padStart(5, '0');

  // 2. TRATAMENTO DOS DADOS DO FORMULÁRIO
  const dataRegistro = new Date();
  const usuarioLogado = Session.getEffectiveUser().getEmail();
  
  // Limpa e extrai a lista de tombos
  const stringTombos = String(payload.tombamentos);
  const arrayTombos = stringTombos
    .split(/[\s,\n]+/) // Quebra por vírgula, espaço ou enter
    .map(t => t.trim())
    .filter(t => t.length > 0);

  if (arrayTombos.length === 0) {
    throw new Error("Nenhum número de tombamento válido foi informado.");
  }

  const tombosFinalFormatado = arrayTombos.join(", ");
  const quantidadeTombos = arrayTombos.length;

  // 3. SALVAR NA PLANILHA (Ordem exata solicitada)
  aba.appendRow([
    seqFormatado,          // SEQUENCIAL DOCUMENTO (10 dígitos)
    dataRegistro,          // DATA_REGISTRO
    usuarioLogado,         // USUARIO
    payload.tipo,          // TIPO
    payload.dataSaida,     // DATA DA SAÍDA
    payload.processo,      // PROCESSO
    quantidadeTombos,      // QUANTIDADE DE TOMBAMENTOS
    tombosFinalFormatado   // TOMBOS
  ]);

  return "OK";
}