function cadastrarNovoCadmat(dados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("CADMAT");
    
    // Se a aba não existir, cria e monta a estrutura até a coluna I
    if (!sheet) {
      sheet = ss.insertSheet("CADMAT");
      const cabecalho = ["CÓDIGO", "DESCRIÇÃO", "", "", "", "", "", "CONTA", "ABREVIATURA ETIQUETA"];
      sheet.appendRow(cabecalho);
      sheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#eaedf1");
    }
    
    // 🌟 Regra de Ouro: Garante os 9 dígitos preenchendo com zeros à esquerda
    const codigoFormatado = String(dados.codigo).trim().padStart(9, '0');
    const descricao = String(dados.descricao).trim().toUpperCase();
    const conta = String(dados.conta).trim();
    const abreviatura = String(dados.abreviatura).trim().toUpperCase();
    
    // Validação básica de segurança antes de salvar
    if (!codigoFormatado || !descricao || !conta || !abreviatura) {
      return { sucesso: false, message: "Todos os campos são obrigatórios." };
    }
    
    // Verifica se o código CADMAT já está cadastrado (evita duplicidade)
    const ultimaLinha = sheet.getLastRow();
    if (ultimaLinha > 1) {
      const codigosExistentes = sheet.getRange(2, 1, ultimaLinha - 1, 1).getValues().flat();
      if (codigosExistentes.includes(codigoFormatado)) {
        return { sucesso: false, message: `O CADMAT ${codigoFormatado} já está cadastrado!` };
      }
    }
    
    // Monta a linha posicionando:
    // Coluna A (1): Código
    // Coluna B (2): Descrição
    // Colunas C-G (3 a 7): Vazias
    // Coluna H (8): Conta
    // Coluna I (9): Abreviatura
    const novaLinha = [codigoFormatado, descricao, "", "", "", "", "", conta, abreviatura];
    sheet.appendRow(novaLinha);
    
    return { sucesso: true, message: `CADMAT ${codigoFormatado} cadastrado com sucesso!` };
    
  } catch (e) {
    return { sucesso: false, message: "Erro no servidor ao cadastrar: " + e.toString() };
  }
}