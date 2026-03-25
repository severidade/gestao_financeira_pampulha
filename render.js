const render = {
  criarTexto(texto, estilo) {
    return SpreadsheetApp.newRichTextValue()
      .setText(String(texto))
      .setTextStyle(estilo)
      .build();
  },

  criarLink(link, estilos) {
    if (link && link.toString().includes("http")) {
      return SpreadsheetApp.newRichTextValue()
        .setText("📱 Ver")
        .setLinkUrl(link)
        .setTextStyle(estilos.link)
        .build();
    }

    return this.criarTexto("-", estilos.normal);
  },

  criarLinkComFallback(link, texto, fallback, estilos) {
    if (link && link.toString().includes("http")) {
      return SpreadsheetApp.newRichTextValue()
        .setText(texto)
        .setLinkUrl(link)
        .setTextStyle(estilos.link)
        .build();
    }

    return this.criarTexto(fallback, estilos.normal);
  },

  criarLinkRecibo(link, estilos) {
    return this.criarLinkComFallback(link, "📄", "🤬", estilos);
  },
};
