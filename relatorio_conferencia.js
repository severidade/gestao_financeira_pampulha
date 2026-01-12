function abrirPainelRelatorio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPassivos = ss.getSheetByName("💸 Passivos_Dados_Brutos");
  
  // --- LÓGICA DE DATAS (Igual à anterior) ---
  const dados = abaPassivos.getDataRange().getValues();
  let opcoesSet = new Set();
  const mapaMeses = { "janeiro":1, "fevereiro":2, "março":3, "abril":4, "maio":5, "junho":6, "julho":7, "agosto":8, "setembro":9, "outubro":10, "novembro":11, "dezembro":12 };

  for (let i = 1; i < dados.length; i++) {
    let mes = dados[i][0];
    let ano = dados[i][1];
    if (mes && ano) opcoesSet.add(`${mes}|${ano}`);
  }

  let listaOpcoes = Array.from(opcoesSet);
  listaOpcoes.sort((a, b) => {
    let [mesA, anoA] = a.split("|");
    let [mesB, anoB] = b.split("|");
    if (anoA !== anoB) return anoB - anoA; 
    return (mapaMeses[mesB.toLowerCase()] || 0) - (mapaMeses[mesA.toLowerCase()] || 0);
  });

  let htmlOptions = listaOpcoes.map(item => {
    let [mes, ano] = item.split("|");
    return `<option value="${mes}|${ano}">${mes.toUpperCase()} / ${ano}</option>`;
  }).join("");

  if (htmlOptions === "") {
    SpreadsheetApp.getUi().alert("Não há dados lançados para gerar relatório.");
    return;
  }

  // --- HTML COM NAVEGAÇÃO DE TELAS ---
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Segoe UI', sans-serif; padding: 0; margin: 0; background-color: #f4f4f4; }
          .container { padding: 20px; }
          
          /* Estilo dos Cards Brancos */
          .card { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin-top: 10px; }
          
          h2 { color: #333; margin-top: 0; text-align: center; }
          label { font-weight: bold; color: #555; display: block; margin-bottom: 5px; }
          
          /* Inputs e Botões */
          select { width: 100%; padding: 12px; margin-bottom: 20px; border-radius: 6px; border: 1px solid #ccc; font-size: 16px; background: #fff; }
          
          .btn { width: 100%; padding: 12px; border: none; border-radius: 6px; font-size: 16px; cursor: pointer; font-weight: bold; transition: 0.3s; margin-bottom: 10px; }
          
          .btn-primary { background-color: #1155cc; color: white; }
          .btn-primary:hover { background-color: #0d47a1; }
          
          .btn-success { background-color: #4CAF50; color: white; }
          .btn-success:hover { background-color: #45a049; }
          
          .btn-secondary { background-color: #fff; color: #555; border: 1px solid #ccc; }
          .btn-secondary:hover { background-color: #eee; }

          /* Animação de Carregamento */
          .loading { display: none; text-align: center; color: #666; margin-top: 20px; }
          .spinner { display: inline-block; width: 20px; height: 20px; border: 3px solid rgba(0,0,0,.1); border-radius: 50%; border-top-color: #1155cc; animation: spin 1s ease-in-out infinite; vertical-align: middle; margin-right: 10px; }
          @keyframes spin { to { transform: rotate(360deg); } }

          /* --- CONTROLE DE TELAS --- */
          #tela-relatorio { display: none; } /* Começa escondida */

          /* --- REGRAS DE IMPRESSÃO --- */
          @media print {
            body { background: white; }
            .card { box-shadow: none; padding: 0; margin: 0; }
            .no-print { display: none !important; } /* Esconde botões na impressão */
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
          // Funcao chamada ao clicar em Visualizar
          function irParaRelatorio() {
            var seletor = document.getElementById("seletorData");
            var valorSelecionado = seletor.value;
            
            // UI: Mostra loading e desabilita botão
            document.getElementById("loading").style.display = "block";
            
            // Backend call
            google.script.run
              .withSuccessHandler(exibirTelaRelatorio)
              .processarRelatorioBackend(valorSelecionado);
          }

          // Callback: Ocorre quando o backend termina
          function exibirTelaRelatorio(htmlRetornado) {
            // Preenche o relatório
            document.getElementById("area-impressao").innerHTML = htmlRetornado;
            
            // Esconde Loading
            document.getElementById("loading").style.display = "none";
            
            // TROCA DE TELAS (Mágica acontece aqui)
            document.getElementById("tela-selecao").style.display = "none";
            document.getElementById("tela-relatorio").style.display = "block";
          }

          // Função do botão "Gerar Outro"
          function voltarParaSelecao() {
            // Limpa o relatório anterior
            document.getElementById("area-impressao").innerHTML = "";
            
            // TROCA DE TELAS INVERSA
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