function obterEstilosPlanilhas() {
  return {
    cabecalho: SpreadsheetApp.newTextStyle()
      .setFontFamily("Jost")
      .setForegroundColor("white")
      .build(),

    normal: SpreadsheetApp.newTextStyle()
      .setFontFamily("Lato")
      .setForegroundColor("black")
      .build(),

    link: SpreadsheetApp.newTextStyle()
      .setUnderline(true)
      .setForegroundColor("#1155cc")
      .build(),

    fundoCabecalho: "black",
  };
}
