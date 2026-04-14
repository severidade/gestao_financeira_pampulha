function atualizarTodasTabelas() {
  const ui = SpreadsheetApp.getUi();

  // ATUALIZA OS BANCOS DE DADOS
  // 1. Passivos
  try {
    atualizarPassivosDadosBrutos();
  } catch (e) {
    ui.alert("Erro Passivos: " + e.message);
    return;
  }

  // 2. Acertos Mensais
  try {
    atualizarAcertosMensaisDadosBrutos();
  } catch (e) {
    ui.alert("Erro Acertos: " + e.message);
    return;
  }

  // ATUALIZA OS DASHBOARDS
  // 3. 💸 Resumo de Despesas e Rateio
  try {
    gerarDashboardRelatorioPassivos();
  } catch (e) {
    ui.alert("Erro Resumo Dashboard: " + e.message);
    return;
  }

  // 4. 🤝 Resumo de Acertos
  try {
    gerarDashboardRelatorioAcertos();
  } catch (e) {
    ui.alert("Erro Links Dashboard: " + e.message);
    return;
  }

  // 6. 💰 Resumo Fluxo de Caixa
  try {
    gerarDashboardFluxoCaixa();
  } catch (e) {
    ui.alert("Erro Links Dashboard: " + e.message);
    return;
  }

  ui.alert("✅ Sucesso! Todas as tabelas e o Dashboard foram atualizados.");
}
