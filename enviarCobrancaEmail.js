function enviarCobrancaEmail(chaveSelecionada, mensagemPersonalizada) {
  const listaDestinatarios = ["oliveira.severo@gmail.com"];
  
  const dados = gerarConteudoEmail(chaveSelecionada, mensagemPersonalizada);

  let inlineImages = {};
  if (dados.qrBlob) {
    inlineImages.qrImagem = dados.qrBlob;
  }

  // 1. Envia o E-mail
  listaDestinatarios.forEach(email => {
    MailApp.sendEmail({
      to: email,
      subject: dados.assunto,
      htmlBody: dados.htmlBody,
      inlineImages: inlineImages,
      attachments: dados.anexos
    });
  });

  // 2. Registra na Planilha (NOVO)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaAcertos = ss.getSheetByName("🤝 Acertos_Mensais_Dados_Brutos");
    const dadosSheet = abaAcertos.getDataRange().getValues();
    const [mesAlvo, anoAlvo] = chaveSelecionada.split("|");
    
    // Data formatada para registro
    const dataHoje = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm");
    const textoStatus = `✅ Enviado em ${dataHoje}`;

    // Procura a linha correta
    for (let i = 1; i < dadosSheet.length; i++) {
      let linhaMes = String(dadosSheet[i][0]).trim();
      let linhaAno = String(dadosSheet[i][1]).trim();
      
      if (linhaMes === mesAlvo && linhaAno === anoAlvo) {
        // Escreve na Coluna G (índice 7, pois getRange começa em 1)
        abaAcertos.getRange(i + 1, 7).setValue(textoStatus);
        break;
      }
    }
  } catch (e) {
    console.log("Erro ao registrar envio na planilha: " + e.message);
  }
}