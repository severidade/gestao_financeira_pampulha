function abrirPainelSelecaoEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAcertos = ss.getSheetByName("🤝 Acertos_Mensais_Dados_Brutos");
  
  if (!abaAcertos) { SpreadsheetApp.getUi().alert("Aba de Acertos não encontrada!"); return; }

  const dados = abaAcertos.getDataRange().getValues();
  let opcoes = [];

  // Começa do 1 para pular cabeçalho
  for (let i = 1; i < dados.length; i++) {
    let mesRef = dados[i][0]; 
    let ano = dados[i][1];
    let vencimento = dados[i][2]; 

    if (mesRef && ano) {
      
      // Filtro de Vencidos
      if (isContaVencida(vencimento)) continue; 

      let dataVencTexto = "-";
      if (vencimento instanceof Date) {
        dataVencTexto = Utilities.formatDate(vencimento, ss.getSpreadsheetTimeZone(), "dd/MM");
      }
      
      let textoOpcao = `${mesRef} / ${ano} (Vence: ${dataVencTexto})`;
      let valorChave = `${mesRef}|${ano}`; 
      
      // MUDANÇA 1: Guardamos mesRef e ano para ordenar depois
      opcoes.push({ 
        html: textoOpcao, 
        val: valorChave,
        rawMes: mesRef, // Dado cru para o Score
        rawAno: ano     // Dado cru para o Score
      });
    }
  }
  
  // MUDANÇA 2: Ordenação Inteligente (Score Maior pro Menor)
  opcoes.sort((a, b) => {
    return ordenarCobrancasPorPeriodo(b.rawMes, b.rawAno) - ordenarCobrancasPorPeriodo(a.rawMes, a.rawAno);
  });

  let htmlOptions = opcoes.map(op => `<option value="${op.val}">${op.html}</option>`).join("");

  if (htmlOptions === "") { 
    SpreadsheetApp.getUi().alert("Não há contas em aberto (dentro do prazo) para enviar."); 
    return; 
  }

  // (O HTML ABAIXO PERMANECE IDÊNTICO AO SEU)
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Segoe UI', sans-serif; padding: 20px; background-color: #f4f4f4; }
          .card { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
          h2 { color: #333; margin-top: 0; text-align: center; font-size: 18px; }
          label { display:block; font-weight:bold; color:#555; margin-bottom:5px;}
          select { width: 100%; padding: 12px; margin: 0 0 20px 0; border-radius: 6px; font-size: 15px; border: 1px solid #ccc; }
          .btn { width: 100%; padding: 12px; border: none; border-radius: 6px; font-size: 16px; cursor: pointer; font-weight: bold; color: white; background-color: #1155cc; transition: 0.2s;}
          .btn:hover { background-color: #0d47a1; }
          .loading { display: none; text-align: center; color: #666; margin-top: 20px; font-size:13px;}
          .spinner { display: inline-block; width: 15px; height: 15px; border: 2px solid rgba(0,0,0,.1); border-radius: 50%; border-top-color: #1155cc; animation: spin 1s infinite; vertical-align: middle; margin-right: 5px; }
          @keyframes spin { to { transform: rotate(360deg); } }
        </style>
      </head>
      <body>
        <div class="card">
          <h2>Enviar Cobrança</h2>
          <label>Selecione a conta:</label>
          <select id="seletor">${htmlOptions}</select>
          <button class="btn" onclick="verPreview()">👁️ Visualizar Antes de Enviar</button>
          <div id="loading" class="loading"><div class="spinner"></div> Carregando pré-visualização...</div>
        </div>
        <script>
          function verPreview() {
            var val = document.getElementById("seletor").value;
            document.getElementById("loading").style.display = "block";
            google.script.run.withSuccessHandler(function() { google.script.host.close(); }).abrirPreviewEmail(val);
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlTemplate).setWidth(420).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Gestão Pampulha');
}