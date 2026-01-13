function abrirPainelRelatorio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPassivos = ss.getSheetByName("💸 Passivos_Dados_Brutos");
  
  // --- LÓGICA DE DATAS ---
  const dados = abaPassivos.getDataRange().getValues();
  let opcoesSet = new Set();
  
  // Mapa robusto para garantir a ordem dos meses
  const mapaMeses = { 
    "janeiro":1, "fevereiro":2, "março":3, "marco":3, // Garante com e sem cedilha
    "abril":4, "maio":5, "junho":6, 
    "julho":7, "agosto":8, "setembro":9, 
    "outubro":10, "novembro":11, "dezembro":12 
  };

  // Coleta os meses e anos disponíveis (Começa do 1 para pular cabeçalho)
  for (let i = 1; i < dados.length; i++) {
    let mes = dados[i][0]; // Coluna A (Mês)
    let ano = dados[i][1]; // Coluna B (Ano)
    
    // Só adiciona se tiver mês E ano preenchidos
    if (mes && ano) {
      // Normaliza para string e remove espaços extras
      let mesStr = String(mes).toLowerCase().trim();
      let anoStr = String(ano).trim();
      opcoesSet.add(`${mesStr}|${anoStr}`);
    }
  }

  let listaOpcoes = Array.from(opcoesSet);

  // --- ORDENAÇÃO INVERSA (Decrescente) ---
  // Lógica: Ano maior vem primeiro. Se o ano for igual, Mês maior vem primeiro.
  listaOpcoes.sort((a, b) => {
    let [mesA, anoA] = a.split("|");
    let [mesB, anoB] = b.split("|");
    
    // Converte ano para número para comparar matemática
    let valAnoA = parseInt(anoA);
    let valAnoB = parseInt(anoB);
    
    // 1. Compara o Ano (Decrescente: B - A)
    if (valAnoA !== valAnoB) {
      return valAnoB - valAnoA; 
    }
    
    // 2. Se o ano for igual, compara o Mês (Decrescente: B - A)
    let valMesA = mapaMeses[mesA] || 0;
    let valMesB = mapaMeses[mesB] || 0;
    
    return valMesB - valMesA;
  });

  // Gera o HTML das opções
  let htmlOptions = listaOpcoes.map(item => {
    let [mes, ano] = item.split("|");
    // Capitaliza a primeira letra do mês para ficar bonito no menu
    let mesFormatado = mes.charAt(0).toUpperCase() + mes.slice(1);
    return `<option value="${mes}|${ano}">${mesFormatado} / ${ano}</option>`;
  }).join("");

  if (htmlOptions === "") {
    SpreadsheetApp.getUi().alert("Não há dados lançados para gerar relatório.");
    return;
  }

  // --- HTML DA JANELA ---
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
            <label for="seletorData">Selecione o período:</label>
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
                <button class="btn btn-form">
                  📝 Lançar Acerto Mensal (Forms)
                </button>
              </a>

              <button class="btn btn-success" onclick="window.print()">
                🖨️ Imprimir / Salvar PDF
              </button>
              
              <button class="btn btn-secondary" onclick="voltarParaSelecao()">
                ⬅️ Gerar Outro Relatório
              </button>
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
              .processarRelatorioBackend(valorSelecionado);
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

  const htmlOutput = HtmlService.createHtmlOutput(htmlTemplate)
      .setWidth(450)
      .setHeight(650);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Gestão Pampulha');
}