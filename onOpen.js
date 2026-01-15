function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('❤️ Gestão Pampulha')
      // --- 1. CADASTROS (Entrada de Dados) ---
      .addItem('💸 Cadastrar Passivo', 'abrirFormularioPassivos')
      .addItem('🤝 Cadastrar Acerto Mensal', 'abrirFormularioAcerto')
      .addItem('💰 Cadastrar Fluxo de Caixa', 'abrirFormularioFluxo')
      .addSeparator()
      
      // --- 2. ATUALIZAÇÕES (Processamento de Dados) ---
      .addItem('💸 Atualizar Tabela Passivos', 'gestao_pampulha_passivos_dados_brutos')
      .addItem('🤝 Atualizar Tabela Acertos Mensais', 'gestao_pampulha_acertos_mensais_dados_brutos')
      .addItem('💰 Atualizar Tabela Fluxo de Caixa', 'gestao_pampulha_fluxo_caixa_dados_brutos')
      .addItem('🔄 Atualizar Todas as Tabelas', 'atualizar_tudo') // <--- Agora está junto aqui
      .addSeparator() 
      
      // --- 3. RELATÓRIOS (Conferência) ---
      .addItem('📋 Passivos Relatório de Conferência', 'abrirPainelRelatorioPassivos')
      .addSeparator() 
      
      // --- 4. AÇÕES FINAIS (Comunicação) ---
      // .addItem('📧 Enviar E-mail (Último Acerto)', 'notificar_novo_acerto')
      .addItem('📧 E-mail Cobrança', 'abrirPainelSelecaoEmail')
      .addToUi();
}

// Função "Mestra" (Mantida a ordem lógica Passivos -> Acertos -> Fluxo)
function atualizar_tudo() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Passivos
  try { gestao_pampulha_passivos_dados_brutos(); } 
  catch (e) { ui.alert("Erro Passivos: " + e.message); return; }

  // 2. Acertos Mensais
  try { gestao_pampulha_acertos_mensais_dados_brutos(); } 
  catch (e) { ui.alert("Erro Acertos: " + e.message); return; }

  // 3. Fluxo de Caixa
  try { gestao_pampulha_fluxo_caixa_dados_brutos(); } 
  catch (e) { ui.alert("Erro Fluxo: " + e.message); return; }

  // 4. Resumo de Conferência (Tabela Esquerda) -> NOME NOVO AQUI
  try { 
    gerarDashboardRelatorioPassivos(); 
  } catch (e) { ui.alert("Erro Resumo Dashboard: " + e.message); return; }

  // 5. Links de Pagamento (Tabela Direita)
  try { 
    gerarDashboardRelatorioAcertos(); 
  } catch (e) { ui.alert("Erro Links Dashboard: " + e.message); return; }

  ui.alert("✅ Sucesso! Todas as tabelas e o Dashboard foram atualizados.");
}

// --- FUNÇÕES DE ABERTURA DOS FORMULÁRIOS ---

function abrirFormularioPassivos() {
  abrirJanelaForms("https://forms.gle/m11gLWD4FZZ1Gykq7", "Cadastrar Passivo");
}

function abrirFormularioAcerto() {
  abrirJanelaForms("https://forms.gle/8NoSATgZ5A9L6Gvw8", "Cadastrar Acerto Mensal");
}

function abrirFormularioFluxo() {
  abrirJanelaForms("https://forms.gle/EELc6Jq3Y71sAco1A", "Cadastrar Fluxo de Caixa");
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