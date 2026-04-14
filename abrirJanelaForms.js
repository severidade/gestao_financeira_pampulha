function abrirJanelaForms(url, titulo) {
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Segoe UI', sans-serif; padding: 20px; text-align: center; background-color: #f4f4f4; }
          .btn { 
            background-color: #673AB7; /* Roxo Forms */
            color: white; 
            padding: 15px 30px; 
            text-decoration: none; 
            border-radius: 8px; 
            font-weight: bold; 
            font-size: 16px; 
            display: inline-block; 
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            transition: 0.3s;
          }
          .btn:hover { background-color: #512DA8; box-shadow: 0 4px 8px rgba(0,0,0,0.3); }
          p { color: #555; margin-bottom: 20px; }
        </style>
      </head>
      <body>
        <p>Clique abaixo para abrir o formulário:</p>
        <a href="${url}" target="_blank" class="btn" onclick="google.script.host.close()">
          📝 ${titulo}
        </a>
      </body>
    </html>
  `;

  const html = HtmlService.createHtmlOutput(htmlTemplate).setWidth(350).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, titulo);
}