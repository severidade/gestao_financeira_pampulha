/**
 * Abre o Painel para Lançar Pagamentos + Upload
 * CORREÇÃO: Tratamento rigoroso de Ano e Valores para evitar erros de gravação.
 */
function abrirPainelBaixaPagamentos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAcertos = ss.getSheetByName("🤝 Acertos_Mensais_Dados_Brutos");
  const abaDB = ss.getSheetByName("💰 Fluxo_Caixa_Dados_Brutos");

  // --- 1. MEMÓRIA DE BLOQUEIO ---
  let chavesExistentes = new Set();
  if (abaDB) {
    const dadosDB = abaDB.getDataRange().getValues();
    for (let i = 1; i < dadosDB.length; i++) {
      let mesDB = String(dadosDB[i][0]).trim();
      let anoDB = String(dadosDB[i][1]).trim();
      if (mesDB && anoDB) chavesExistentes.add(`${mesDB}|${anoDB}`);
    }
  }

  // --- 2. COLETAR OPÇÕES ---
  const dados = abaAcertos.getDataRange().getValues();
  let opcoes = [];

  for (let i = 1; i < dados.length; i++) {
    // Garante que são strings limpas
    let mesRef = String(dados[i][0]).trim();
    let ano = String(dados[i][1]).trim();

    // TRATAMENTO DE VALOR (Blindagem)
    // Remove R$, espaços e converte para número puro para o JS
    let valorRaw = dados[i][3];
    let valorNumero = 0;

    if (typeof valorRaw === "number") {
      valorNumero = valorRaw;
    } else {
      // Se for texto "R$ 1.200,50", limpa tudo
      let vStr = String(valorRaw).replace("R$", "").trim();
      if (vStr.includes(",") && vStr.includes(".")) {
        vStr = vStr.replace(/\./g, "").replace(",", "."); // 1.000,00 -> 1000.00
      } else if (vStr.includes(",")) {
        vStr = vStr.replace(",", "."); // 50,00 -> 50.00
      }
      valorNumero = parseFloat(vStr) || 0;
    }

    if (mesRef && ano) {
      if (chavesExistentes.has(`${mesRef}|${ano}`)) continue;

      opcoes.push({
        rotulo: `${mesRef} / ${ano}`,
        valor: valorNumero, // Envia número puro (float)
        mes: mesRef,
        ano: ano,
      });
    }
  }

  if (opcoes.length === 0) {
    SpreadsheetApp.getUi().alert(
      "Todas as cobranças geradas já constam no Fluxo de Caixa! ✅",
    );
    return;
  }

  // --- 3. ORDENAÇÃO ---
  opcoes.sort((a, b) => {
    let indA = 0,
      indB = 0;
    let mesA = a.mes,
      mesB = b.mes;
    let matchA = a.mes.match(/\((\d+)\)/);
    if (matchA) {
      indA = matchA[1];
      mesA = a.mes.split("(")[0].trim();
    }
    let matchB = b.mes.match(/\((\d+)\)/);
    if (matchB) {
      indB = matchB[1];
      mesB = b.mes.split("(")[0].trim();
    }
    return (
      ordenarCobrancasPorPeriodo(mesB, b.ano, indB) -
      ordenarCobrancasPorPeriodo(mesA, a.ano, indA)
    );
  });

  const opcoesJSON = JSON.stringify(opcoes);

  // --- 4. HTML ---
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
          select, input[type="date"], input[type="text"] { width: 100%; padding: 10px; margin-bottom: 10px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; font-size: 14px; }
          input[type="file"] { font-size: 12px; margin-bottom: 15px; width: 100%; }
          .row-pessoa { display: flex; flex-direction: column; padding: 10px 0; border-bottom: 1px solid #eee; }
          .row-header { display: flex; align-items: center; gap: 10px; margin-bottom: 8px; }
          .row-body { display: flex; gap: 10px; }
          .col-nome { flex-grow: 1; font-weight: bold; }
          .col-data, .col-valor { width: 50%; }
          .btn { width: 100%; padding: 12px; border: none; border-radius: 6px; font-size: 16px; cursor: pointer; font-weight: bold; color: white; background-color: #1e8e3e; }
          .btn:disabled { background-color: #ccc; }
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
          <h2>💰 Lançar Recebimento</h2>
          <label>Selecione a Cobrança:</label>
          <select id="seletorCobranca" onchange="atualizarValoresSugeridos()">
            <option value="" disabled selected>-- Escolha o Mês/Ciclo --</option>
          </select>
          <div style="text-align: right;"><span id="displayValorBase" class="badge-valor">Valor Base: R$ 0,00</span></div>
        </div>

        <div class="card">
          <h2>Quem Pagou?</h2>
          ${gerarHtmlPessoa("Marco")}
          ${gerarHtmlPessoa("Janaina")}
          ${gerarHtmlPessoa("Adriana")}
        </div>

        <button class="btn" id="btnSalvar" onclick="iniciarSalvamento()">💾 Salvar e Upload</button>

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
            ['Marco', 'Janaina', 'Adriana'].forEach(n => document.getElementById("data"+n).value = hoje);
          }

          function atualizarValoresSugeridos() {
            let idx = document.getElementById("seletorCobranca").value;
            if (idx === "") return;
            let item = listaCobrancas[idx];
            // Formata para o input visual (ex: 150,00)
            let valorFmt = item.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2, maximumFractionDigits: 2});
            document.getElementById("displayValorBase").innerText = "Cota: R$ " + valorFmt;
            ['valorMarco', 'valorJanaina', 'valorAdriana'].forEach(id => document.getElementById(id).value = valorFmt);
          }

          function togglePessoa(nome) {
            let check = document.getElementById("check" + nome).checked;
            document.getElementById("data" + nome).disabled = !check;
            document.getElementById("valor" + nome).disabled = !check;
            document.getElementById("file" + nome).disabled = !check;
          }

          function lerArquivo(file) {
            return new Promise((resolve, reject) => {
              const reader = new FileReader();
              reader.onload = () => resolve({
                nome: file.name, tipo: file.type,
                base64: reader.result.split(',')[1]
              });
              reader.onerror = error => reject(error);
              reader.readAsDataURL(file);
            });
          }

          async function iniciarSalvamento() {
            let idx = document.getElementById("seletorCobranca").value;
            if (idx === "") { alert("Selecione um mês!"); return; }
            let item = listaCobrancas[idx];

            // Verifica ANTES de enviar se o ano e mês estão ok
            if (!item.ano || !item.mes) {
                alert("Erro nos dados da cobrança (Ano ou Mês vazio). Verifique a planilha de Acertos.");
                return;
            }

            let btn = document.getElementById("btnSalvar");
            btn.innerText = "Processando...";
            btn.disabled = true;

            let pagamentos = [];
            let pessoas = ['Marco', 'Janaina', 'Adriana'];

            for (let nome of pessoas) {
              if (document.getElementById("check" + nome).checked) {
                let inputArquivo = document.getElementById("file" + nome);
                let dadosArquivo = null;

                if (inputArquivo.files.length > 0) {
                  try {
                    dadosArquivo = await lerArquivo(inputArquivo.files[0]);
                  } catch (e) {
                    alert("Erro ao ler arquivo de " + nome);
                    btn.disabled = false; return;
                  }
                }

                pagamentos.push({
                   pessoa: nome,
                   data: document.getElementById("data" + nome).value,
                   valor: document.getElementById("valor" + nome).value, // Envia o texto como está no input
                   arquivo: dadosArquivo
                });
              }
            }

            let payload = {
              mesRef: item.mes,
              anoRef: item.ano,
              pagamentos: pagamentos
            };

            google.script.run
              .withSuccessHandler(() => {
                 alert("Sucesso! Pagamentos salvos.");
                 google.script.host.close();
              })
              .withFailureHandler((e) => {
                 alert("Erro no script: " + e);
                 btn.innerText = "Tentar Novamente";
                 btn.disabled = false;
              })
              .processarEntradaFluxoCaixa(payload);
          }
        </script>
      </body>
    </html>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlTemplate)
    .setWidth(420)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(
    htmlOutput,
    "Fluxo de Caixa + Recibos",
  );
}

function gerarHtmlPessoa(nome) {
  return `
    <div class="row-pessoa">
      <div class="row-header">
        <label class="switch"><input type="checkbox" id="check${nome}" onchange="togglePessoa('${nome}')"><span class="slider"></span></label>
        <span class="col-nome">${nome === "Janaina" ? "Janaína" : nome}</span>
      </div>
      <div class="row-body">
         <div class="col-data"><input type="date" id="data${nome}" disabled></div>
         <div class="col-valor"><input type="text" id="valor${nome}" placeholder="0,00" disabled></div>
      </div>
      <div style="margin-top:5px;">
         <input type="file" id="file${nome}" accept="application/pdf,image/*" disabled>
      </div>
    </div>
  `;
}
