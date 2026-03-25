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

  // Alinhamento coluna
  abaDestino.getRange(1, 1, numLinhas, 1).setHorizontalAlignment("left"); // Mês
  abaDestino.getRange(1, 2, numLinhas, 2).setHorizontalAlignment("center"); // Ano + Vencimento
  abaDestino.getRange(1, 4, numLinhas, 1).setHorizontalAlignment("right"); // Valor
  abaDestino.getRange(1, 5, numLinhas, 1).setHorizontalAlignment("center"); // QR
  abaDestino.getRange(1, 6, numLinhas, 1).setHorizontalAlignment("left"); // Pix
  abaDestino.getRange(1, 7, numLinhas, 1).setHorizontalAlignment("center"); // Status

  // Alinhamento vertical
  abaDestino.getRange(1, 1, numLinhas, 7).setVerticalAlignment("middle"); // Toda a Tabela

  // Largura coluna
  abaDestino.setColumnWidth(6, 600); // Codigo pix

  // WRAP
  abaDestino.getRange(1, 6, numLinhas, 1).setWrap(true); // Chave Pix
  abaDestino.getRange(1, 7, numLinhas, 1).setWrap(true); // Status
}
