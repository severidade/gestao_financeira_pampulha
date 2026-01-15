function abrirPainelRelatorioPassivos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPassivos = ss.getSheetByName("💸 Passivos_Dados_Brutos");
  
  // --- LÓGICA DE DADOS ---
  const dados = abaPassivos.getDataRange().getValues();
  let opcoesSet = new Set();
  
  // Coleta os meses, anos e tipos disponíveis
  for (let i = 1; i < dados.length; i++) {
    let mesCru = dados[i][0]; // Coluna A (Ex: "Janeiro (1)")
    let ano = dados[i][1];    // Coluna B
    
    if (mesCru && ano) {
      let mesStr = String(mesCru).trim();
      let anoStr = String(ano).trim();
      
      // --- CORREÇÃO: Extrair o índice de dentro do texto do mês ---
      let indice = 0;
      let mesLimpo = mesStr.toLowerCase();

      // Procura por (número) no nome. Ex: "janeiro (2)"
      const match = mesStr.match(/\((\d+)\)/);
      
      if (match) {
        indice = parseInt(match[1]); // Pega o 2
        // Remove o "(2)" do nome para ficar só "janeiro"
        mesLimpo = mesStr.replace(/\s*\(\d+\)/, "").trim().toLowerCase();
      }
      // ------------------------------------------------------------

      opcoesSet.add(`${mesLimpo}|${anoStr}|${indice}`);
    }
  }

  let listaOpcoes = Array.from(opcoesSet);

  // --- ORDENAÇÃO: DECRESCENTE (Extra 2 -> Extra 1 -> Padrão) ---
  listaOpcoes.sort((a, b) => {
    let [mesA, anoA, indA] = a.split("|");
    let [mesB, anoB, indB] = b.split("|");
    
    // b - a garante que o MAIOR índice e data apareçam primeiro
    return ordenarCobrancasPorPeriodo(mesB, anoB, indB) - ordenarCobrancasPorPeriodo(mesA, anoA, indA);
  });

  // --- MONTAGEM DO TEXTO (FORMATO LIMPO) ---
  let htmlOptions = listaOpcoes.map(item => {
    let [mes, ano, indice] = item.split("|");
    
    let mesFormatado = mes.charAt(0).toUpperCase() + mes.slice(1);
    
    // Lógica de Exibição
    let textoVisivel = "";
    
    if (indice == "0") {
        // Se for padrão (0), mostra apenas "Janeiro 2026"
        textoVisivel = `${mesFormatado} ${ano}`;
    } else {
        // Se for extra, mostra "Janeiro 2026 - Extra X"
        textoVisivel = `${mesFormatado} ${ano} - Extra ${indice}`;
    }

    return `<option value="${mes}|${ano}|${indice}">${textoVisivel}</option>`;
  }).join("");

  if (htmlOptions === "") {
    SpreadsheetApp.getUi().alert("Não há dados lançados para gerar relatório.");
    return;
  }

  // --- HTML DA JANELA (MANTIDO IGUAL) ---
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Segoe UI', sans-serif; padding: 0; margin: 0; background-color: #f4f4f4; }
          .container { padding: 20px; }
          .card { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin-top: 10px; }
          h2 { color: #333; margin-top: 0; text-align: center; }
          label { font-weight: bold; color: #555; display: block; margin-bottom: 5px; }
          select { width: 100%; padding: 12px; margin-bottom: 20px; border-radius: 6px; border: 1px solid #ccc; font-size: 16px; background: #fff; }
          .btn { width: 100%; padding: 12px; border: none; border-radius: 6px; font-size: 16px; cursor: pointer; font-weight: bold; transition: 0.3s; margin-bottom: 10px; }
          .btn-primary { background-color: #1155cc; color: white; }
          .btn-primary:hover { background-color: #0d47a1; }
          .btn-form { background-color: #673AB7; color: white; }
          .btn-form:hover { background-color: #512DA8; }
          .btn-success { background-color: #4CAF50; color: white; }
          .btn-success:hover { background-color: #45a049; }
          .btn-secondary { background-color: #fff; color: #555; border: 1px solid #ccc; }
          .btn-secondary:hover { background-color: #eee; }
          .loading { display: none; text-align: center; color: #666; margin-top: 20px; }
          .spinner { display: inline-block; width: 20px; height: 20px; border: 3px solid rgba(0,0,0,.1); border-radius: 50%; border-top-color: #1155cc; animation: spin 1s ease-in-out infinite; vertical-align: middle; margin-right: 10px; }
          @keyframes spin { to { transform: rotate(360deg); } }
          #tela-relatorio { display: none; } 
          @media print {
            body { background: white; }
            .card { box-shadow: none; padding: 0; margin: 0; }
            .no-print { display: none !important; } 
            #area-impressao { display: block !important; }
          }
        </style>
      </head>
      <body>
        
        <div id="tela-selecao" class="container">
          <div class="card">
            <h2>📊 Conferência Mensal</h2>
            <br>
            <label for="seletorData">Selecione o Ciclo de Cobrança:</label>
            <select id="seletorData">
              ${htmlOptions}
            </select>
            
            <button class="btn btn-primary" onclick="irParaRelatorio()">
              🔍 Visualizar Relatório
            </button>
            
            <div id="loading" class="loading">
              <div class="spinner"></div> Gerando relatório...
            </div>
          </div>
        </div>

        <div id="tela-relatorio" class="container">
          <div class="card">
            <div id="area-impressao"></div>
            <div class="no-print" style="margin-top: 25px; border-top: 1px solid #eee; padding-top: 15px;">
              <a href="https://forms.gle/4TF91DSR91EGyTKv8" target="_blank" style="text-decoration:none;">
                <button class="btn btn-form">📝 Lançar Acerto Mensal (Forms)</button>
              </a>
              <button class="btn btn-success" onclick="window.print()">🖨️ Imprimir / Salvar PDF</button>
              <button class="btn btn-secondary" onclick="voltarParaSelecao()">⬅️ Gerar Outro Relatório</button>
            </div>
          </div>
        </div>

        <script>
          function irParaRelatorio() {
            var seletor = document.getElementById("seletorData");
            var valorSelecionado = seletor.value;
            document.getElementById("loading").style.display = "block";
            google.script.run
              .withSuccessHandler(exibirTelaRelatorio)
              .processarRelatorioPassivos(valorSelecionado);
          }

          function exibirTelaRelatorio(htmlRetornado) {
            document.getElementById("area-impressao").innerHTML = htmlRetornado;
            document.getElementById("loading").style.display = "none";
            document.getElementById("tela-selecao").style.display = "none";
            document.getElementById("tela-relatorio").style.display = "block";
          }

          function voltarParaSelecao() {
            document.getElementById("area-impressao").innerHTML = "";
            document.getElementById("tela-relatorio").style.display = "none";
            document.getElementById("tela-selecao").style.display = "block";
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlTemplate).setWidth(450).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Gestão Pampulha');
}