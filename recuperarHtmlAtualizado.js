function recuperarHtmlAtualizado(chaveSelecionada, mensagemPersonalizada) {
  const dados = gerarConteudoEmail(chaveSelecionada, mensagemPersonalizada);
  
  // Converte QR Blob para Base64 para exibir no HTML do Modal
  let base64Qr = "";
  if (dados.qrBlob) {
    base64Qr = Utilities.base64Encode(dados.qrBlob.getBytes());
  }
  
  // Substitui o CID pelo Base64 para visualização no navegador
  let htmlPronto = dados.htmlBody.replace('src="cid:qrImagem"', `src="data:image/png;base64,${base64Qr}"`);
  
  return htmlPronto;
}