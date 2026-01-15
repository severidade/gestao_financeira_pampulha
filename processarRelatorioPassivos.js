function processarRelatorioPassivos(chaveData) {
  // --- PROTEÇÃO CONTRA ERRO ---
  if (!chaveData) chaveData = "janeiro|2026|0"; 

  // Recebe os dados limpos vindos do menu
  const [mesSolicitado, anoSolicitado, indiceSolicitado] = chaveData.split("|");
  
  const NUMERO_PESSOAS = 3; 
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPassivos = ss.getSheetByName("💸 Passivos_Dados_Brutos");

  // --- FORMATAÇÃO ---
  const fmt = (v) => parseFloat(v).toLocaleString('pt-BR', {
    style: 'currency', 
    currency: 'BRL', 
    minimumFractionDigits: 2, 
    maximumFractionDigits: 2 
  });

  const tratarNumero = (v) => {
    if (typeof v === 'number') return v;
    if (!v) return 0;
    return parseFloat(String(v).replace(/[^\d,-]/g, '').replace(',', '.')) || 0;
  };

  // 2. LER DADOS
  const dadosPassivos = abaPassivos.getDataRange().getValues();
  let listaHtml = "";
  let somaTotalPassivos = 0;

  for (let i = 1; i < dadosPassivos.length; i++) {
    
    // Dados crus da linha
    let linhaMesCru = String(dadosPassivos[i][0]).trim(); // Ex: "Janeiro (1)"
    let linhaAno = String(dadosPassivos[i][1]).trim();
    
    // --- LÓGICA DE EXTRAÇÃO (IGUAL AO MENU) ---
    // Precisamos separar o nome "Janeiro" do índice "(1)"
    let linhaMesLimpo = linhaMesCru.toLowerCase();
    let linhaIndice = "0"; // Assume padrão (0)

    const match = linhaMesCru.match(/\((\d+)\)/);
    if (match) {
      linhaIndice = match[1]; // Pega o número (ex: "1")
      linhaMesLimpo = linhaMesCru.replace(/\s*\(\d+\)/, "").trim().toLowerCase(); // Vira "janeiro"
    }
    // -------------------------------------------

    // 4. COMPARAÇÃO
    // Agora comparamos maçãs com maçãs (dados limpos com dados limpos)
    if (linhaMesLimpo === String(mesSolicitado).toLowerCase() && 
        linhaAno === String(anoSolicitado) &&
        linhaIndice === String(indiceSolicitado)) {
      
      let servicoNome = dadosPassivos[i][2]; // Nome do serviço
      let valorItem = tratarNumero(dadosPassivos[i][3]); // Valor
      somaTotalPassivos += valorItem;
      
      listaHtml += `<div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee; padding:5px 0;">
                      <span>${servicoNome}</span>
                      <strong>${fmt(valorItem)}</strong>
                    </div>`;
    }
  }

  // ADICIONA TOTAL
  if (somaTotalPassivos > 0) {
    listaHtml += `
      <div style="display:flex; justify-content:space-between; border-top:2px solid #555; margin-top:5px; padding-top:5px; color:#000;">
        <span style="font-weight:bold;">TOTAL</span>
        <span style="font-weight:bold;">${fmt(somaTotalPassivos)}</span>
      </div>
    `;
  }

  // CÁLCULO RATEIO
  let valorRateioCalculado = somaTotalPassivos / NUMERO_PESSOAS;

  // TÍTULO DINÂMICO
  let textoTipo = (indiceSolicitado === "0") ? "COBRANÇA PADRÃO" : `COBRANÇA EXTRA ${indiceSolicitado}`;
  let corTitulo = (indiceSolicitado === "0") ? "#1155cc" : "#1155cc"; // coloquei a mesma cor 

  // --- HTML FINAL ---
  let htmlFinal = `
    <div style="text-align: center; margin-bottom: 20px;">
      <h2 style="color:${corTitulo}; margin:0; font-size: 22px;">RELATÓRIO DE CONFERÊNCIA</h2>
      <h3 style="color:#555; margin:5px 0 0 0; font-weight:normal;">${mesSolicitado.toUpperCase()} / ${anoSolicitado}</h3>
      <small style="color:#888; font-weight:bold;">${textoTipo}</small>
    </div>
    
    <div style="background:#f0f4ff; padding:15px; border:1px solid #cce0ff; border-radius:8px; margin-bottom:20px; text-align:center;">
      <small style="text-transform: uppercase; color: #555; font-size: 11px;">Valor da cota individual</small><br>
      <span style="font-size:28px; font-weight:bold; color:#1155cc;">
        ${fmt(valorRateioCalculado)}
      </span>
      
      <div style="font-size:11px; color:#777; margin-top:5px;">
        (${fmt(somaTotalPassivos)} dividido por ${NUMERO_PESSOAS})
      </div>

    </div>

    <div style="margin-bottom:10px; font-weight:bold; border-bottom: 2px solid ${corTitulo}; padding-bottom: 5px; color:#333;">
      COMPOSIÇÃO DAS CONTAS
    </div>
    
    <div style="font-size: 14px; min-height: 80px; max-height:300px; overflow-y:auto;">
      ${listaHtml || "<p style='text-align:center; color:#999; margin-top:20px;'><em>Nenhuma despesa encontrada para este agrupamento.</em></p>"}
    </div>
  `;

  return htmlFinal;
}