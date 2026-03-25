function escreverTabelaAcertoMensal(
  abaDestino,
  dados,
  valoresNumericos,
  estilos,
) {
  const numLinhas = dados.length;

  abaDestino.getRange(1, 1, numLinhas, 7).setRichTextValues(dados);

  // cabeçalho
  abaDestino
    .getRange(1, 1, 1, 7)
    .setBackground(estilos.fundoCabecalho)
    .setFontWeight("bold");

  abaDestino.getRange(1, 4, numLinhas, 1).setNumberFormat("R$ #,##0.00");

  if (valoresNumericos.length > 0) {
    abaDestino
      .getRange(1, 4, valoresNumericos.length, 1)
      .setValues(valoresNumericos);
  }

  // Formatação
  abaDestino.getRange(1, 1, numLinhas, 7).setHorizontalAlignment("left"); // Toda Tabela
  abaDestino.getRange(1, 1, numLinhas, 7).setVerticalAlignment("middle"); // Toda a Tabela
  abaDestino.getRange(1, 5, numLinhas, 1).setHorizontalAlignment("center"); // QR Code
  abaDestino.setColumnWidth(6, 600); //chave pix
  abaDestino.getRange(1, 6, numLinhas, 1).setWrap(true); // chave pix (wrap)
  abaDestino.autoResizeColumns(1, 7);
}
