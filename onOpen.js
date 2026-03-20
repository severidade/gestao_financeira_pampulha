function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("❤️ Gestão Sítio Boa Vista")
    // --- 1. CADASTROS (Entrada de Dados) ---
    .addItem("💸 Cadastrar Passivo", "abrirFormularioPassivos")
    .addItem("🤝 Cadastrar Acerto Mensal", "abrirFormularioAcerto")
    .addSeparator()

    // --- 2. ATUALIZAÇÕES (Processamento de Dados) ---
    .addItem("💸 Atualizar Tabela Passivos", "atualizarPassivosDadosBrutos")
    .addItem(
      "🤝 Atualizar Tabela Acertos Mensais - verificar",
      "atualizarAcertosMensaisDadosBrutos",
    )
    .addItem("🔄 Atualizar Todas as Tabelas", "atualizarTodasTabelas")
    .addSeparator()

    // --- 3. RELATÓRIOS (Conferência) ---
    .addItem(
      "📋 Passivos Relatório de Conferência",
      "abrirPainelRelatorioPassivos",
    )
    .addSeparator()

    // --- 4. AÇÕES FINAIS (Comunicação) ---
    .addItem("📧 E-mail Cobrança", "abrirPainelSelecaoEmail")
    // .addSeparator()

    // --- 5. AÇÕES FINAIS (Comunicação) ---
    // .addItem('💰 Fluxo Caixa Lancamento', 'abrirPainelBaixaPagamentos')
    .addToUi();
}

// --- FUNÇÕES DE ABERTURA DOS FORMULÁRIOS ---
function abrirFormularioPassivos() {
  abrirJanelaForms("https://forms.gle/Zsksrqiz1UG5AQaG8", "Cadastrar Passivo");
}

function abrirFormularioAcerto() {
  abrirJanelaForms(
    "https://forms.gle/8xTp4VAx4UN79Fe36",
    "Cadastrar Acerto Mensal",
  );
}
