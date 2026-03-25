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

  abaDestino.getRange(1, 1, numLinhas, 7).setHorizontalAlignment("center");
  abaDestino.getRange(1, 6, numLinhas, 1).setHorizontalAlignment("left");

  abaDestino.autoResizeColumns(1, 7);
}
