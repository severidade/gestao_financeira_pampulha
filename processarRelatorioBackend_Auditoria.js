function processarRelatorioBackend(chaveData) {
  // --- PROTEÇÃO CONTRA ERRO (Caso rode pelo editor) ---
  if (!chaveData) chaveData = "janeiro|2026"; 

  const [mesSolicitado, anoSolicitado] = chaveData.split("|");
  
  // --- CONFIGURAÇÃO DO RATEIO ---
  const NUMERO_PESSOAS = 3; 
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPassivos = ss.getSheetByName("💸 Passivos_Dados_Brutos");
  const abaAcertos = ss.getSheetByName("🤝 Acertos_Mensais_Dados_Brutos");

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

  // --- NOVO: ADICIONA A LINHA DE TOTAL AO FINAL DA LISTA ---
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

  // 2. BUSCA O ACERTO E O QR CODE (COM RICHTEXT PARA PEGAR LINK)
  const ultimaLinha = abaAcertos.getLastRow();
  const rangeAcertos = abaAcertos.getRange(2, 1, ultimaLinha - 1, 6);
  const dadosAcertos = rangeAcertos.getValues();
  const formulasAcertos = rangeAcertos.getFormulas();
  const richTextAcertos = rangeAcertos.getRichTextValues(); 

  let valorCobradoReal = 0;
  let acertoEncontrado = false;
  let idQrCode = "";
  
  for (let i = 0; i < dadosAcertos.length; i++) {
    if (String(dadosAcertos[i][0]).toLowerCase() === String(mesSolicitado).toLowerCase() && 
        String(dadosAcertos[i][1]) === anoSolicitado) {
      
      valorCobradoReal = tratarNumero(dadosAcertos[i][2]);
      acertoEncontrado = true;

      // Extração do QR Code
      const cellRichText = richTextAcertos[i][3];
      const urlLink = cellRichText ? cellRichText.getLinkUrl() : null;
      const formula = formulasAcertos[i][3]; 
      const valorTexto = dadosAcertos[i][3]; 

      if (urlLink && urlLink.includes("id=")) {
        idQrCode = urlLink.split("id=")[1];
      } else if (formula && formula.includes("id=")) {
        let match = formula.match(/id=([a-zA-Z0-9_-]+)/);
        if (match) idQrCode = match[1];
      } else if (typeof valorTexto === 'string' && valorTexto.includes("id=")) {
        idQrCode = valorTexto.split("id=")[1];
      }
      break;
    }
  }

  // 3. PROCESSAMENTO DA IMAGEM
  let imgTag = "";
  if (idQrCode) {
    try {
      let arquivo = DriveApp.getFileById(idQrCode);
      let blob = arquivo.getBlob();
      let base64 = Utilities.base64Encode(blob.getBytes());
      let tipo = blob.getContentType();
      imgTag = `<img src="data:${tipo};base64,${base64}" style="width:140px; height:140px; object-fit:contain; border:1px solid #ccc; border-radius:5px; margin-bottom:10px;">`;
    } catch (e) {
      imgTag = `<div style="color:red; font-size:11px;">Erro Imagem</div>`;
    }
  } else {
    imgTag = `<div style="color:#aaa; font-size:11px; padding:20px; border:1px dashed #ccc;">Sem QR Code</div>`;
  }

  // 4. LÓGICA DE CONFERÊNCIA
  let diferenca = valorCobradoReal - valorRateioCalculado;
  diferenca = Math.round(diferenca * 100) / 100;

  let htmlConferencia = "";
  if (!acertoEncontrado) {
     htmlConferencia = `<div style="color:orange; font-weight:bold;">⚠️ Acerto Mensal não lançado.</div>`;
  } else if (Math.abs(diferenca) < 0.05) { 
     htmlConferencia = `<div style="color:green; font-weight:bold; font-size:16px;">✅ TUDO CERTO! Valor Confere.</div>`;
  } else {
     let cor = "red"; 
     let texto = diferenca > 0 ? "a MAIOR" : "a MENOR";
     htmlConferencia = `
       <div style="color:${cor}; font-weight:bold; font-size:16px;">⚠️ DIVERGÊNCIA DE VALOR</div>
       <div style="color:${cor}; font-size:14px; margin-top:5px;">
         Cobrado: <strong>${fmt(valorCobradoReal)}</strong> | Ideal: <strong>${fmt(valorRateioCalculado)}</strong>
         <br>Dif: ${fmt(Math.abs(diferenca))} ${texto}.
       </div>
     `;
  }

  // 5. HTML FINAL
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
        ( valor dividido ${NUMERO_PESSOAS} pessoas)
      </div>
    </div>

    <div style="margin-bottom:10px; font-weight:bold; border-bottom: 2px solid #1155cc; padding-bottom: 5px; color:#333;">
      COMPOSIÇÃO DAS CONTAS
    </div>
    <div style="font-size: 14px; min-height: 80px; max-height:160px; overflow-y:auto;">
      ${listaHtml || "<p style='text-align:center; color:#999; margin-top:20px;'><em>Nenhuma despesa encontrada.</em></p>"}
    </div>
    
    <div style="margin-top:25px; background-color: #f9f9f9; border: 1px solid #ddd; border-radius: 5px; padding: 15px; text-align:center;">
      
      <div style="font-weight:bold; color:#333; margin-bottom:10px; border-bottom:1px solid #ccc; padding-bottom:5px;">
        AUDITORIA DE COBRANÇA
      </div>
      
      ${imgTag}

      <div style="display:flex; justify-content:space-between; margin: 10px 0; font-size: 14px;">
        <span>Valor Lançado (Acertos):</span>
        <strong>${acertoEncontrado ? fmt(valorCobradoReal) : "---"}</strong>
      </div>
      
      <div style="border-top: 1px solid #ccc; padding-top: 10px; text-align: center;">
        ${htmlConferencia}
      </div>

    </div>
  `;

  return htmlFinal;
}