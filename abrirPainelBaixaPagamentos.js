/**
 * Abre o Painel para Lançar Pagamentos recebidos (Fluxo de Caixa)
 * FILTRO NOVO: Não exibe meses que já possuem linha criada no DB.
 */
function abrirPainelBaixaPagamentos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAcertos = ss.getSheetByName("🤝 Acertos_Mensais_Dados_Brutos");
  const abaDB = ss.getSheetByName("📥 DB_Pagamentos"); // Aba de verificação
  
  // --- 1. MEMÓRIA: Mapear o que JÁ EXISTE no Fluxo de Caixa ---
  let chavesExistentes = new Set();
  
  if (abaDB) {
    const dadosDB = abaDB.getDataRange().getValues();
    // Começa do 1 para pular cabeçalho
    for (let i = 1; i < dadosDB.length; i++) {
      let mesDB = String(dadosDB[i][0]).trim(); // Coluna A (Mês Ref)
      let anoDB = String(dadosDB[i][1]).trim(); // Coluna B (Ano)
      
      // Se tem mês e ano, adiciona na lista negra
      if (mesDB && anoDB) {
        chavesExistentes.add(`${mesDB}|${anoDB}`);
      }
    }
  }

  // --- 2. COLETAR OPÇÕES (Filtrando as existentes) ---
  const dados = abaAcertos.getDataRange().getValues();
  let opcoes = [];

  for (let i = 1; i < dados.length; i++) {
    let mesRef = String(dados[i][0]).trim(); 
    let ano = String(dados[i][1]).trim();
    let valorCota = dados[i][3]; // Valor da cota individual
    
    if (mesRef && ano) {
      
      // A REGRA DE BLOQUEIO:
      // Verifica se essa combinação "Mês|Ano" já está no DB. Se estiver, pula.
      if (chavesExistentes.has(`${mesRef}|${ano}`)) {
        continue; 
      }

      let textoOpcao = `${mesRef} / ${ano}`;
      
      opcoes.push({
        rotulo: textoOpcao,
        valor: valorCota,
        mes: mesRef,
        ano: ano
      });
    }
  }

  // Se não houver nada pendente, avisa e para.
  if (opcoes.length === 0) {
    SpreadsheetApp.getUi().alert("Todas as cobranças geradas já constam no Fluxo de Caixa! ✅");
    return;
  }

  // --- 3. ORDENAÇÃO (Mantida) ---
  opcoes.sort((a, b) => {
    let indA = 0, indB = 0;
    let mesA = a.mes, mesB = b.mes;

    let matchA = a.mes.match(/\((\d+)\)/);
    if(matchA) { indA = matchA[1]; mesA = a.mes.split("(")[0].trim(); }
    
    let matchB = b.mes.match(/\((\d+)\)/);
    if(matchB) { indB = matchB[1]; mesB = b.mes.split("(")[0].trim(); }

    // Função externa de ordenação
    return ordenarCobrancasPorPeriodo(mesB, b.ano, indB) - ordenarCobrancasPorPeriodo(mesA, a.ano, indA);
  });

  const opcoesJSON = JSON.stringify(opcoes);

  // --- 4. HTML (Mantido igual) ---
  const htmlTemplate = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">
        <style>
          body { font-family: 'Roboto', sans-serif; background-color: #f0f2f5; margin: 0; padding: 20px; color: #333; }
          .card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); margin-bottom: 15px; }
          h2 { color: #1a73e8; margin-top: 0; font-size: 18px; display: flex; align-items: center; gap: 8px; }
          
          label { display: block; font-weight: 500; margin-bottom: 5px; font-size: 14px; color: #555; }
          select, input[type="date"], input[type="text"] { width: 100%; padding: 10px; margin-bottom: 15px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; font-size: 14px; }
          
          .row-pessoa { display: flex; align-items: center; gap: 10px; padding: 10px 0; border-bottom: 1px solid #eee; }
          .row-pessoa:last-child { border-bottom: none; }
          .col-check { width: 30px; }
          .col-nome { flex-grow: 1; font-weight: bold; }
          .col-data { width: 130px; }
          .col-valor { width: 100px; }

          .btn { width: 100%; padding: 12px; border: none; border-radius: 6px; font-size: 16px; cursor: pointer; font-weight: bold; transition: 0.2s; color: white; background-color: #1e8e3e; }
          .btn:hover { background-color: #187033; }
          
          .switch { position: relative; display: inline-block; width: 34px; height: 20px; }
          .switch input { opacity: 0; width: 0; height: 0; }
          .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: #ccc; transition: .4s; border-radius: 34px; }
          .slider:before { position: absolute; content: ""; height: 14px; width: 14px; left: 3px; bottom: 3px; background-color: white; transition: .4s; border-radius: 50%; }
          input:checked + .slider { background-color: #1a73e8; }
          input:checked + .slider:before { transform: translateX(14px); }

          .badge-valor { background: #e8f0fe; color: #1967d2; padding: 4px 8px; border-radius: 12px; font-size: 12px; font-weight: bold; }
        </style>
      </head>
      <body>
        
        <div class="card">
          <h2>💰 Lançar Recebimento (Novo)</h2>
          
          <label>Selecione a Cobrança Disponível:</label>
          <select id="seletorCobranca" onchange="atualizarValoresSugeridos()">
            <option value="" disabled selected>-- Escolha o Mês/Ciclo --</option>
          </select>
          
          <div style="margin-top: -10px; margin-bottom: 15px; text-align: right;">
            <span id="displayValorBase" class="badge-valor">Valor Base: R$ 0,00</span>
          </div>
        </div>

        <div class="card">
          <h2>Quem Pagou?</h2>
          
          <div class="row-pessoa">
            <div class="col-check">
               <label class="switch"><input type="checkbox" id="checkMarco" onchange="togglePessoa('Marco')"><span class="slider"></span></label>
            </div>
            <div class="col-nome">Marco</div>
            <div class="col-data">
               <input type="date" id="dataMarco" disabled>
            </div>
            <div class="col-valor">
               <input type="text" id="valorMarco" placeholder="R$ 0,00" disabled>
            </div>
          </div>

          <div class="row-pessoa">
            <div class="col-check">
               <label class="switch"><input type="checkbox" id="checkJanaina" onchange="togglePessoa('Janaina')"><span class="slider"></span></label>
            </div>
            <div class="col-nome">Janaína</div>
            <div class="col-data">
               <input type="date" id="dataJanaina" disabled>
            </div>
            <div class="col-valor">
               <input type="text" id="valorJanaina" placeholder="R$ 0,00" disabled>
            </div>
          </div>

          <div class="row-pessoa">
            <div class="col-check">
               <label class="switch"><input type="checkbox" id="checkAdriana" onchange="togglePessoa('Adriana')"><span class="slider"></span></label>
            </div>
            <div class="col-nome">Adriana</div>
            <div class="col-data">
               <input type="date" id="dataAdriana" disabled>
            </div>
            <div class="col-valor">
               <input type="text" id="valorAdriana" placeholder="R$ 0,00" disabled>
            </div>
          </div>

        </div>

        <button class="btn" id="btnSalvar" onclick="salvar()">💾 Criar Registro no Fluxo</button>

        <script>
          let listaCobrancas = ${opcoesJSON};
          
          window.onload = function() {
            let select = document.getElementById("seletorCobranca");
            listaCobrancas.forEach((item, index) => {
               let opt = document.createElement("option");
               opt.value = index; 
               opt.text = item.rotulo;
               select.add(opt);
            });
            
            let hoje = new Date().toISOString().split('T')[0];
            document.getElementById("dataMarco").value = hoje;
            document.getElementById("dataJanaina").value = hoje;
            document.getElementById("dataAdriana").value = hoje;
          }

          function atualizarValoresSugeridos() {
            let idx = document.getElementById("seletorCobranca").value;
            if (idx === "") return;
            
            let item = listaCobrancas[idx];
            let valorFmt = parseFloat(item.valor).toLocaleString('pt-BR', {minimumFractionDigits: 2});
            document.getElementById("displayValorBase").innerText = "Cota Individual: R$ " + valorFmt;
            
            ['valorMarco', 'valorJanaina', 'valorAdriana'].forEach(id => {
               document.getElementById(id).value = valorFmt;
            });
          }

          function togglePessoa(nome) {
            let check = document.getElementById("check" + nome).checked;
            document.getElementById("data" + nome).disabled = !check;
            document.getElementById("valor" + nome).disabled = !check;
          }

          function salvar() {
            let idx = document.getElementById("seletorCobranca").value;
            if (idx === "") { alert("Selecione um mês de referência!"); return; }
            
            let item = listaCobrancas[idx];
            let payload = {
              mesRef: item.mes,
              anoRef: item.ano,
              pagamentos: []
            };

            ['Marco', 'Janaina', 'Adriana'].forEach(nome => {
              if (document.getElementById("check" + nome).checked) {
                payload.pagamentos.push({
                   pessoa: nome,
                   data: document.getElementById("data" + nome).value,
                   valor: document.getElementById("valor" + nome).value
                });
              }
            });

            // Permite salvar mesmo sem ninguém marcado (cria a linha zerada)
            // Se preferir obrigar pelo menos um, descomente abaixo:
            /*
            if (payload.pagamentos.length === 0) {
               let confirmar = confirm("Ninguém foi marcado como pago. Deseja apenas criar a linha vazia?");
               if (!confirmar) return;
            }
            */

            let btn = document.getElementById("btnSalvar");
            btn.innerText = "Salvando...";
            btn.disabled = true;

            google.script.run
              .withSuccessHandler(function() {
                 alert("Registro criado com sucesso! Este mês sairá da lista de pendentes.");
                 google.script.host.close();
              })
              .withFailureHandler(function(e) {
                 alert("Erro: " + e);
                 btn.innerText = "Tentar Novamente";
                 btn.disabled = false;
              })
              .processarEntradaFluxoCaixa(payload);
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlTemplate).setWidth(400).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Lançar no Fluxo de Caixa');
}