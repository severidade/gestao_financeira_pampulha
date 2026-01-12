function processarRelatorioBackend(chaveData) {
  // --- PROTEÇÃO CONTRA ERRO (Caso rode pelo editor) ---
  if (!chaveData) chaveData = "janeiro|2026"; 

  const [mesSolicitado, anoSolicitado] = chaveData.split("|");
  
  // --- CONFIGURAÇÃO DO RATEIO ---
  const NUMERO_PESSOAS = 3; 
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPassivos = ss.getSheetByName("💸 Passivos_Dados_Brutos");
  // A aba de Acertos não é mais necessária para este relatório

  const fmt = (v) => parseFloat(v).toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});
  const tratarNumero = (v) => {
    if (typeof v === 'number') return v;
    if (!v) return 0;
    return parseFloat(String(v).replace(/[^\d,-]/g, '').replace(',', '.')) || 0;
  };

  // 1. CALCULA O TOTAL DE PASSIVOS E GERA A LISTA
  const dadosPassivos = abaPassivos.getDataRange().getValues();
  let listaHtml = "";
  let somaTotalPassivos = 0;

  for (let i = 1; i < dadosPassivos.length; i++) {
    if (String(dadosPassivos[i][0]).toLowerCase() === String(mesSolicitado).toLowerCase() && 
        String(dadosPassivos[i][1]) === anoSolicitado) {
      
      let valorItem = tratarNumero(dadosPassivos[i][3]);
      somaTotalPassivos += valorItem;
      
      listaHtml += `<div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee; padding:5px 0;">
                      <span>${dadosPassivos[i][2]}</span>
                      <strong>${fmt(valorItem)}</strong>
                    </div>`;
    }
  }

  // ADICIONA A LINHA DE TOTAL AO FINAL DA LISTA
  if (somaTotalPassivos > 0) {
    listaHtml += `
      <div style="display:flex; justify-content:space-between; border-top:2px solid #555; margin-top:5px; padding-top:5px; color:#000;">
        <span style="font-weight:bold;">TOTAL</span>
        <span style="font-weight:bold;">${fmt(somaTotalPassivos)}</span>
      </div>
    `;
  }

  // CÁLCULO DO RATEIO INDIVIDUAL
  let valorRateioCalculado = somaTotalPassivos / NUMERO_PESSOAS;

  // --- HTML FINAL (Sem o bloco de Auditoria/QR Code) ---
  let htmlFinal = `
    <div style="text-align: center; margin-bottom: 20px;">
      <h2 style="color:#1155cc; margin:0; font-size: 22px;">RELATÓRIO DE CONFERÊNCIA</h2>
      <h3 style="color:#555; margin:5px 0 0 0; font-weight:normal;">${mesSolicitado.toUpperCase()} / ${anoSolicitado}</h3>
    </div>
    
    <div style="background:#f0f4ff; padding:15px; border:1px solid #cce0ff; border-radius:8px; margin-bottom:20px; text-align:center;">
      <small style="text-transform: uppercase; color: #555; font-size: 11px;">Valor da cota individual</small><br>
      <span style="font-size:28px; font-weight:bold; color:#1155cc;">
        ${fmt(valorRateioCalculado)}
      </span>
      <div style="font-size:11px; color:#777; margin-top:5px;">
        (valor dividido por ${NUMERO_PESSOAS} pessoas)
      </div>
    </div>

    <div style="margin-bottom:10px; font-weight:bold; border-bottom: 2px solid #1155cc; padding-bottom: 5px; color:#333;">
      COMPOSIÇÃO DAS CONTAS
    </div>
    <div style="font-size: 14px; min-height: 80px; max-height:250px; overflow-y:auto;">
      ${listaHtml || "<p style='text-align:center; color:#999; margin-top:20px;'><em>Nenhuma despesa encontrada.</em></p>"}
    </div>
  `;

  return htmlFinal;
}