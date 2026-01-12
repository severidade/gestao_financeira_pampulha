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

  // --- ESTILOS ---
  const estiloNormal = SpreadsheetApp.newTextStyle().setUnderline(false).setForegroundColor("black").build();
  const estiloLink = SpreadsheetApp.newTextStyle().setUnderline(true).setForegroundColor("#1155cc").build();

  // --- FUNÇÕES AJUDANTES ---
  function tratarValor(valorStr) {
    if (!valorStr) return 0;
    let numero = parseFloat(valorStr.toString().replace(',', '.').replace(/[^\d.-]/g, ''));
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

  // --- PROCESSAMENTO ---
  
  // 1. LIMPEZA IMEDIATA: Garante que apaga dados antigos antes de qualquer checagem
  abaDestino.clear();

  const dadosOrigem = abaOrigem.getDataRange().getDisplayValues(); 
  
  // Prepara os arrays de saída
  const saidaRichText = [];
  const valoresNumericos = [ ["Valor"] ]; // Título da coluna numérica
  
  // Monta o Cabeçalho (Isso garante que a tabela sempre tenha títulos, mesmo vazia)
  const cabecalho = ["Mês Ref.", "Ano", "Valor", "QR Code", "Chave Pix"].map(txt => 
    SpreadsheetApp.newRichTextValue().setText(txt).setTextStyle(estiloNormal).build()
  );
  saidaRichText.push(cabecalho);

  // 2. SÓ PROCESSA SE TIVER DADOS (Mais de 1 linha)
  if (dadosOrigem.length >= 2) {
    
    let linhasDados = dadosOrigem.slice(1);

    // ORDENAÇÃO
    linhasDados.sort(function(a, b) {
      const anoA = parseInt(a[2]) || 0;
      const anoB = parseInt(b[2]) || 0;
      const mesA = obterNumeroMes(a[1]);
      const mesB = obterNumeroMes(b[1]);

      if (anoA !== anoB) return anoA - anoB; 
      return mesA - mesB;   
    });

    linhasDados.forEach(linha => {
      // 0:Timestamp | 1:Mês | 2:Ano | 3:QR Code | 4:Chave | 5:Valor
      const mesRef = linha[1];
      const anoRef = linha[2];
      const linkQr = linha[3];
      const chavePix = linha[4];
      const valorBruto = linha[5];

      if (mesRef || valorBruto) {
        let valorNumerico = tratarValor(valorBruto); 

        // QR Code
        let rtQr;
        if (linkQr && linkQr.toString().includes("http")) {
          rtQr = SpreadsheetApp.newRichTextValue()
            .setText("📱 Abrir")
            .setLinkUrl(linkQr)
            .setTextStyle(estiloLink)
            .build();
        } else {
          rtQr = SpreadsheetApp.newRichTextValue().setText("-").setTextStyle(estiloNormal).build();
        }

        // Chave Pix
        let rtChave = SpreadsheetApp.newRichTextValue().setText(chavePix || "-").setTextStyle(estiloNormal).build();
        let rtMes = SpreadsheetApp.newRichTextValue().setText(mesRef).setTextStyle(estiloNormal).build();
        let rtAno = SpreadsheetApp.newRichTextValue().setText(anoRef).setTextStyle(estiloNormal).build();
        
        // Valor Texto
        let textoValor = valorNumerico.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        let rtValor = SpreadsheetApp.newRichTextValue().setText(textoValor).setTextStyle(estiloNormal).build(); 

        saidaRichText.push([rtMes, rtAno, rtValor, rtQr, rtChave]);
        valoresNumericos.push([valorNumerico]);
      }
    });
  }

  // --- ESCREVER NA PLANILHA ---
  // Escreve o resultado (Seja só o cabeçalho ou o cabeçalho + dados)
  if (saidaRichText.length > 0) {
    const numLinhas = saidaRichText.length;
    
    // Formata Coluna Valor
    abaDestino.getRange(1, 3, numLinhas, 1).setNumberFormat("R$ #,##0.00");
    
    // Escreve RichText
    abaDestino.getRange(1, 1, numLinhas, 5).setRichTextValues(saidaRichText);
    
    // Sobrescreve números reais na coluna Valor
    if(valoresNumericos.length > 0){
       abaDestino.getRange(1, 3, valoresNumericos.length, 1).setValues(valoresNumericos);
    }
    
    abaDestino.autoResizeColumns(1, 5);
  }
}