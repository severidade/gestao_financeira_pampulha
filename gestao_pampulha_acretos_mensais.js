function gestao_pampulha_acertos_mensais_dados_brutos() {
  // --- CONFIGURAÇÕES ---
  const idPlanilhaOrigem = "1e8MOW4QLmR89zoQFul70Y6deUpPGCcz6MgXJuc3upfc"; 
  const nomeAbaOrigem = "Respostas ao formulário 1"; 
  const nomeAbaDestino = "🤝 Acertos_Mensais_Dados_Brutos"; 

  // --- ACESSAR PLANILHAS ---
  const ssOrigem = SpreadsheetApp.openById(idPlanilhaOrigem);
  const abaOrigem = ssOrigem.getSheetByName(nomeAbaOrigem);
  
  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  let abaDestino = ssDestino.getSheetByName(nomeAbaDestino);

  if (!abaDestino) {
    abaDestino = ssDestino.insertSheet(nomeAbaDestino);
  }

  // ============================================================
  // 🛡️ PASSO 1: ATIVAR A MEMÓRIA (ANTES DE APAGAR)
  // ============================================================
  const memoriaStatus = {};
  
  // Pega todos os dados que estão na planilha AGORA
  const dadosAtuais = abaDestino.getDataRange().getValues();
  
  // Se tiver dados (mais que 1 linha), vamos memorizar
  if (dadosAtuais.length > 1) {
    for (let i = 1; i < dadosAtuais.length; i++) {
      let mesChave = String(dadosAtuais[i][0]).trim(); // Coluna A (Ex: Janeiro (1))
      let anoChave = String(dadosAtuais[i][1]).trim(); // Coluna B (Ex: 2026)
      let status = dadosAtuais[i][6]; // Coluna G (Onde você escreveu ✅ Pago)
      
      // Se tiver algo escrito na Coluna G, guarda no "bolso" do script
      if (mesChave && anoChave && status !== "") {
        let chaveUnica = `${mesChave}|${anoChave}`;
        memoriaStatus[chaveUnica] = status;
      }
    }
  }
  // ============================================================

  // --- ESTILOS ---
  const estiloNormal = SpreadsheetApp.newTextStyle().setUnderline(false).setForegroundColor("black").build();
  const estiloLink = SpreadsheetApp.newTextStyle().setUnderline(true).setForegroundColor("#1155cc").build(); 

  function tratarValor(valorStr) {
    if (!valorStr) return 0;
    let v = valorStr.toString().trim();
    if (!isNaN(v) && !v.includes(',') && !v.includes('.')) return parseFloat(v);
    if (v.includes(',')) v = v.replace("R$", "").replace(/\./g, "").replace(",", ".");
    else if (v.includes('.')) v = v.replace("R$", "");
    let numero = parseFloat(v);
    return isNaN(numero) ? 0 : numero;
  }

  function obterNumeroMes(nomeMes) {
    if (!nomeMes) return 0;
    const mes = nomeMes.toString().trim().toLowerCase();
    const mapa = {
      "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3,
      "abril": 4, "maio": 5, "junho": 6,
      "julho": 7, "agosto": 8, "setembro": 9,
      "outubro": 10, "novembro": 11, "dezembro": 12
    };
    return mapa[mes] || 0;
  }

  // --- PASSO 2: APAGAR TUDO (Agora é seguro, pois já memorizamos) ---
  abaDestino.clear();
  SpreadsheetApp.flush();

  const dadosOrigem = abaOrigem.getDataRange().getDisplayValues(); 
  
  const saidaRichText = [];
  const valoresNumericos = [ ["Valor"] ]; 
  
  const titulos = ["Mês Ref.", "Ano", "Vencimento", "Valor", "QR Code", "Chave Pix", "Status Envio"];
  
  const cabecalho = titulos.map(txt => 
    SpreadsheetApp.newRichTextValue().setText(txt).setTextStyle(estiloNormal).build()
  );
  saidaRichText.push(cabecalho);

  if (dadosOrigem.length >= 2) {
    let linhasDados = dadosOrigem.slice(1);

    linhasDados.sort(function(a, b) {
      const anoA = parseInt(a[2]) || 0;
      const anoB = parseInt(b[2]) || 0;
      if (anoA !== anoB) return anoA - anoB; 
      
      const mesA = obterNumeroMes(a[1]);
      const mesB = obterNumeroMes(b[1]);
      if (mesA !== mesB) return mesA - mesB;

      const supA = parseInt(a[6]) || 0;
      const supB = parseInt(b[6]) || 0;
      return supA - supB; 
    });

    linhasDados.forEach(linha => {
      const mesRef = linha[1];
      const anoRef = linha[2];
      const linkQrOriginal = linha[3]; 
      const chavePix = linha[4]; 
      const valorBruto = linha[5];
      const inputSuplementar = linha[6]; 
      const dataVencimento = linha[7];    

      if (mesRef || valorBruto) {
        let valorNumerico = tratarValor(valorBruto); 

        let textoMesComposto = mesRef;
        let indiceSup = 0;
        if (inputSuplementar && inputSuplementar.toString().trim() !== "") {
            indiceSup = parseInt(inputSuplementar.toString().trim());
        }
        if (indiceSup > 0) {
            textoMesComposto = `${mesRef} (${indiceSup})`;
        }

        let rtQr;
        if (linkQrOriginal && linkQrOriginal.toString().includes("http")) {
          rtQr = SpreadsheetApp.newRichTextValue()
            .setText("📱 Abrir")
            .setLinkUrl(linkQrOriginal)
            .setTextStyle(estiloLink)
            .build();
        } else {
          rtQr = SpreadsheetApp.newRichTextValue().setText("-").setTextStyle(estiloNormal).build();
        }

        // ========================================================
        // 🛡️ PASSO 3: RESTAURAR O STATUS
        // ========================================================
        let statusParaGravar = "-";
        let chaveAtual = `${textoMesComposto}|${anoRef}`;
        
        // Verifica se temos algo guardado para este Mês/Ano
        if (memoriaStatus[chaveAtual]) {
          statusParaGravar = memoriaStatus[chaveAtual];
        }
        // ========================================================

        let textoVencimento = dataVencimento || "-";

        let rtMes = SpreadsheetApp.newRichTextValue().setText(textoMesComposto).setTextStyle(estiloNormal).build();
        let rtAno = SpreadsheetApp.newRichTextValue().setText(anoRef).setTextStyle(estiloNormal).build();
        let rtVencimento = SpreadsheetApp.newRichTextValue().setText(textoVencimento).setTextStyle(estiloNormal).build();
        let rtChave = SpreadsheetApp.newRichTextValue().setText(chavePix || "-").setTextStyle(estiloNormal).build();
        let textoValor = valorNumerico.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        let rtValor = SpreadsheetApp.newRichTextValue().setText(textoValor).setTextStyle(estiloNormal).build(); 
        
        // Coluna G: Grava o que recuperamos da memória
        let rtStatus = SpreadsheetApp.newRichTextValue().setText(statusParaGravar).setTextStyle(estiloNormal).build();

        saidaRichText.push([rtMes, rtAno, rtVencimento, rtValor, rtQr, rtChave, rtStatus]);
        valoresNumericos.push([valorNumerico]);
      }
    });
  }

  // --- ESCREVER NA PLANILHA ---
  if (saidaRichText.length > 0) {
    const numLinhas = saidaRichText.length;
    
    abaDestino.getRange(1, 1, numLinhas, 7).setRichTextValues(saidaRichText);
    abaDestino.getRange(1, 4, numLinhas, 1).setNumberFormat("R$ #,##0.00");
    
    if(valoresNumericos.length > 0){
       abaDestino.getRange(1, 4, valoresNumericos.length, 1).setValues(valoresNumericos);
    }
    
    abaDestino.getRange(1, 1, numLinhas, 7).setHorizontalAlignment("center");
    abaDestino.getRange(1, 6, numLinhas, 1).setHorizontalAlignment("left");
    
    abaDestino.autoResizeColumns(1, 7);
  }
}