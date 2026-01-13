function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('❤️ Gestão Pampulha')
      // --- BLOCO DE CADASTROS (FORMS) ---
      .addItem('💸 Cadastrar Passivo', 'abrirFormularioPassivos')
      .addItem('💰 Cadastrar Fluxo de Caixa', 'abrirFormularioFluxo')
      .addItem('🤝 Cadastrar Acerto Mensal', 'abrirFormularioAcerto')
      .addSeparator()
      
      // --- BLOCO DE ATUALIZAÇÃO INDIVIDUAL ---
      .addItem('💸 Atualizar Tabela Passivos', 'gestao_pampulha_passivos_dados_brutos')
      .addItem('💰 Atualizar Tabela Fluxo de Caixa', 'gestao_pampulha_fluxo_caixa_dados_brutos')
      .addItem('🤝 Atualizar Tabela Acertos Mensais', 'gestao_pampulha_acertos_mensais_dados_brutos')
      .addSeparator() 
      
      // --- BLOCO GERAL ---
      .addItem('🔄 Atualizar TUDO', 'atualizar_tudo')
      .addSeparator() 
      .addItem('📋 Relatório de Conferência', 'abrirPainelRelatorio')
      .addSeparator() 
      .addItem('📧 Enviar E-mail (Último Acerto)', 'notificar_novo_acerto')
      .addToUi();
}

// Função "Mestra" que roda as três funções em sequência com tratamento de erro detalhado
function atualizar_tudo() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Roda Passivos
  try {
    gestao_pampulha_passivos_dados_brutos(); 
  } catch (e) {
    ui.alert("Erro ao atualizar Passivos: " + e.message);
    return; // Para o script se der erro aqui
  }

  // 2. Roda Fluxo de Caixa
  try {
    gestao_pampulha_fluxo_caixa_dados_brutos(); 
  } catch (e) {
    ui.alert("Erro ao atualizar Fluxo de Caixa: " + e.message);
    return; // Para o script se der erro aqui
  }

  // 3. Roda Acertos Mensais
  try {
    gestao_pampulha_acertos_mensais_dados_brutos(); 
  } catch (e) {
    ui.alert("Erro ao atualizar Acertos Mensais: " + e.message);
    return; // Para o script se der erro aqui
  }

  // 4. Avisa que acabou
  ui.alert("✅ Sucesso! Todas as tabelas foram atualizadas.");
}

// --- FUNÇÕES DE ABERTURA DOS FORMULÁRIOS ---

function abrirFormularioPassivos() {
  abrirJanelaForms("https://forms.gle/m11gLWD4FZZ1Gykq7", "Cadastrar Passivo");
}

function abrirFormularioFluxo() {
  abrirJanelaForms("https://forms.gle/EELc6Jq3Y71sAco1A", "Cadastrar Fluxo de Caixa");
}

function abrirFormularioAcerto() {
  abrirJanelaForms("https://forms.gle/8NoSATgZ5A9L6Gvw8", "Cadastrar Acerto Mensal");
}

// --- FUNÇÃO AUXILIAR PARA GERAR A JANELA HTML ---
function abrirJanelaForms(url, titulo) {
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Segoe UI', sans-serif; padding: 20px; text-align: center; background-color: #f4f4f4; }
          .btn { 
            background-color: #673AB7; /* Roxo Forms */
            color: white; 
            padding: 15px 30px; 
            text-decoration: none; 
            border-radius: 8px; 
            font-weight: bold; 
            font-size: 16px; 
            display: inline-block; 
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            transition: 0.3s;
          }
          .btn:hover { background-color: #512DA8; box-shadow: 0 4px 8px rgba(0,0,0,0.3); }
          p { color: #555; margin-bottom: 20px; }
        </style>
      </head>
      <body>
        <p>Clique abaixo para abrir o formulário:</p>
        <a href="${url}" target="_blank" class="btn" onclick="google.script.host.close()">
          📝 ${titulo}
        </a>
      </body>
    </html>
  `;

  const html = HtmlService.createHtmlOutput(htmlTemplate).setWidth(350).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, titulo);
}