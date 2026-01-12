function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('❤️ Gestão Pampulha')
      .addItem('💸 Atualizar Passivos', 'gestao_pampulha_passivos_dados_brutos')
      .addItem('💰 Atualizar Fluxo de Caixa', 'gestao_pampulha_fluxo_caixa_dados_brutos')
      .addItem('🤝 Atualizar Acertos Mensais', 'gestao_pampulha_acertos_mensais_dados_brutos') // <--- NOVO ITEM
      .addSeparator() 
      .addItem('🔄 Atualizar TUDO', 'atualizar_tudo')
      .addSeparator() 
      .addItem('📋 relatorio', 'abrirPainelRelatorio')
      .addSeparator() 
      .addItem('📧 Enviar E-mail (Último Acerto)', 'notificar_novo_acerto')
      .addToUi();
}

// Função "Mestra" que roda as três funções em sequência
function atualizar_tudo() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Roda Passivos
  try {
    gestao_pampulha_passivos_dados_brutos(); 
  } catch (e) {
    ui.alert("Erro ao atualizar Passivos: " + e.message);
    return; 
  }

  // 2. Roda Fluxo de Caixa
  try {
    gestao_pampulha_fluxo_caixa_dados_brutos(); 
  } catch (e) {
    ui.alert("Erro ao atualizar Fluxo de Caixa: " + e.message);
    return;
  }

  // 3. Roda Acertos Mensais (NOVO)
  try {
    gestao_pampulha_acertos_mensais_dados_brutos(); 
  } catch (e) {
    ui.alert("Erro ao atualizar Acertos Mensais: " + e.message);
    return;
  }

  // 4. Avisa que acabou
  ui.alert("✅ Sucesso! Todas as tabelas foram atualizadas.");
}