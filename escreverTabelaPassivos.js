function escreverTabelaPassivos(abaDestino, dados, valoresNumericos, estilos) {
  const numLinhas = dados.length;

  abaDestino.getRange(1, 1, numLinhas, 6).setRichTextValues(dados);

  // cabeçalho
  abaDestino
    .getRange(1, 1, 1, 6)
    .setBackground(estilos.fundoCabecalho)
    .setFontWeight("bold");

  // valores
  abaDestino.getRange(1, 4, numLinhas, 1).setNumberFormat("R$ #,##0.00");

  // if (valoresNumericos.length > 0) {
  //   abaDestino
  //     .getRange(1, 4, valoresNumericos.length, 1)
  //     .setValues(valoresNumericos);
  // }

  if (valoresNumericos.length > 0) {
    abaDestino
      .getRange(2, 4, valoresNumericos.length, 1)
      .setValues(valoresNumericos);
  }

  //alinhamento
  abaDestino.getRange(1, 1, numLinhas, 6).setHorizontalAlignment("left");
  abaDestino.getRange(1, 4, numLinhas, 1).setHorizontalAlignment("right");
  abaDestino.getRange(1, 5, numLinhas, 1).setHorizontalAlignment("center");

  abaDestino.autoResizeColumns(1, 6);
}
