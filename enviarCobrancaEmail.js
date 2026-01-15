function enviarCobrancaEmail(chaveSelecionada, mensagemPersonalizada) {
  const listaDestinatarios = ["oliveira.severo@gmail.com"];
  
  // Passamos a mensagem personalizada aqui
  const dados = gerarConteudoEmail(chaveSelecionada, mensagemPersonalizada);

  let inlineImages = {};
  if (dados.qrBlob) {
    inlineImages.qrImagem = dados.qrBlob;
  }

  listaDestinatarios.forEach(email => {
    MailApp.sendEmail({
      to: email,
      subject: dados.assunto,
      htmlBody: dados.htmlBody,
      inlineImages: inlineImages,
      attachments: dados.anexos
    });
  });
}