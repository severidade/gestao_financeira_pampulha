function enviarCobrancaEmail(chaveSelecionada, mensagemPersonalizada) {
  const listaDestinatarios = [
    "familia.pimentelbhz@gmail.com",
    "pimenteljanaina@gmail.com",
  ];
  
  const dados = gerarConteudoEmail(chaveSelecionada, mensagemPersonalizada);

  let inlineImages = {};
  if (dados.qrBlob) {
    inlineImages.qrImagem = dados.qrBlob;
  }

  // 1. Envia UM ÚNICO e-mail com todos em cópia
  MailApp.sendEmail({
    to: listaDestinatarios.join(","),
    cc: "alessandramarabh@gmail.com", // Alessandra copiada
    subject: dados.assunto,
    htmlBody: dados.htmlBody,
    inlineImages: inlineImages,
    attachments: dados.anexos
  });

  // 2. Registra na Planilha
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaAcertos = ss.getSheetByName("🤝 Acertos_Mensais_Dados_Brutos");
    const dadosSheet = abaAcertos.getDataRange().getValues();
    const [mesAlvo, anoAlvo] = chaveSelecionada.split("|");
    
    const dataHoje = Utilities.formatDate(
      new Date(),
      ss.getSpreadsheetTimeZone(),
      "dd/MM/yyyy HH:mm"
    );

    const textoStatus = `✅ Enviado em ${dataHoje}`;

    // Procura a linha correta
    for (let i = 1; i < dadosSheet.length; i++) {
      let linhaMes = String(dadosSheet[i][0]).trim();
      let linhaAno = String(dadosSheet[i][1]).trim();
      
      if (linhaMes === mesAlvo && linhaAno === anoAlvo) {
        abaAcertos.getRange(i + 1, 7).setValue(textoStatus);
        break;
      }
    }

  } catch (e) {
    console.log("Erro ao registrar envio na planilha: " + e.message);
  }
}