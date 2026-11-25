// Retorna a lista completa de usuários para o painel de gerenciamento (Apenas ADMIN)
function listarUsuariosPainel() {
  if (getPerfilUsuario() !== "ADMIN") {
    throw new Error("Acesso negado. Apenas administradores podem gerenciar usuários.");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("PERFIS");
  if (!aba) return [];
  
  const dados = aba.getDataRange().getValues();
  if (dados.length <= 1) return [];
  
  return dados.slice(1).map(linha => ({
    email: String(linha[0]).trim(),
    perfil: String(linha[1]).trim() === "OPERADOR" ? "DBM" : String(linha[1]).trim(), // Converte legados de OPERADOR para DBM na visualização
    nome: String(linha[2]).trim()
  }));
}

// Salva ou atualiza um usuário (Apenas ADMIN)
function salvarUsuarioPainel(usuario) {
  if (getPerfilUsuario() !== "ADMIN") {
    throw new Error("Acesso negado. Apenas administradores podem gerenciar usuários.");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("PERFIS");
  const dados = aba.getDataRange().getValues();
  
  const emailBusca = String(usuario.email).trim().toLowerCase();
  let perfilSalvar = usuario.perfil.toUpperCase();
  if (perfilSalvar === "OPERADOR") perfilSalvar = "DBM"; // Garante consistência
  
  let linhaIndex = -1;
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).trim().toLowerCase() === emailBusca) {
      linhaIndex = i + 1;
      break;
    }
  }
  
  if (linhaIndex !== -1) {
    aba.getRange(linhaIndex, 2).setValue(perfilSalvar);
    aba.getRange(linhaIndex, 3).setValue(usuario.nome);
  } else {
    aba.appendRow([emailBusca, perfilSalvar, usuario.nome]);
  }
  
  return "OK";
}

// Exclui um usuário da planilha (Apenas ADMIN)
function excluirUsuarioPainel(email) {
  if (getPerfilUsuario() !== "ADMIN") {
    throw new Error("Acesso negado. Apenas administradores podem gerenciar usuários.");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("PERFIS");
  const dados = aba.getDataRange().getValues();
  
  const emailBusca = String(email).trim().toLowerCase();
  
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).trim().toLowerCase() === emailBusca) {
      aba.deleteRow(i + 1);
      return "OK";
    }
  }
  
  throw new Error("Usuário não localizado para exclusão.");
}

// Retorna o perfil do usuário logado atualmente (Garante tratamento para DBM)
function getPerfilUsuario() {
  try {
    const emailLogado = Session.getEffectiveUser().getEmail().trim().toLowerCase();
    const dados = getDadosAba("PERFIS");
    
    if (!dados || dados.length === 0) return "VISUALIZADOR";
    
    const linha = dados.find(l => String(l[0]).trim().toLowerCase() === emailLogado);
    if (!linha) return "VISUALIZADOR";
    
    let perfil = String(linha[1]).trim().toUpperCase();
    if (perfil === "OPERADOR") perfil = "DBM"; // Se estiver escrito OPERADOR na planilha, assume DBM
    
    return perfil;
  } catch(e) {
    console.error("Erro em getPerfilUsuario: " + e.message);
    return "VISUALIZADOR";
  }
}

// Função auxiliar para obter os dados de uma aba
function getDadosAba(nomeAba) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(nomeAba);
  if (!aba) return [];
  return aba.getDataRange().getValues();
}