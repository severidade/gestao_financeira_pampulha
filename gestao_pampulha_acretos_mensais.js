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
  // Removi o estiloNegrito pois não será mais usado

  // --- FUNÇÕES AJUDANTES ---
function tratarValor(valorStr) {
    if (!valorStr) return 0;
    
    // Converte para string para garantir
    let v = valorStr.toString().trim();
    
    // Se for apenas número (ex: 100), retorna direto
    if (!isNaN(v) && !v.includes(',') && !v.includes('.')) {
        return parseFloat(v);
    }

    // LÓGICA DE DETECÇÃO:
    // Se tiver vírgula, assumimos que é decimal PT-BR (133,33)
    if (v.includes(',')) {
        // Remove R$ e pontos de milhar, troca vírgula por ponto
        v = v.replace("R$", "").replace(/\./g, "").replace(",", ".");
    } 
    // Se tiver ponto e NÃO tiver vírgula, assumimos que é decimal EN (133.33)
    else if (v.includes('.')) {
        // Remove R$ apenas (mantém o ponto como decimal)
        v = v.replace("R$", "");
    }

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

  // --- PROCESSAMENTO ---
  
  // 1. LIMPEZA
  abaDestino.clear();
  SpreadsheetApp.flush();

  const dadosOrigem = abaOrigem.getDataRange().getDisplayValues(); 
  
  const saidaRichText = [];
  const valoresNumericos = [ ["Valor"] ]; 
  
  // 2. CABEÇALHO (6 Colunas)
  const titulos = ["Mês Ref.", "Ano", "Vencimento", "Valor", "QR Code", "Chave Pix"];
  
  // MUDANÇA AQUI: Usa estiloNormal em vez de estiloNegrito
  const cabecalho = titulos.map(txt => 
    SpreadsheetApp.newRichTextValue().setText(txt).setTextStyle(estiloNormal).build()
  );
  saidaRichText.push(cabecalho);

  // 3. PROCESSA DADOS
  if (dadosOrigem.length >= 2) {
    
    let linhasDados = dadosOrigem.slice(1);

    // ORDENAÇÃO: Ano -> Mês -> Suplementar
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
      // MAPEAMENTO:
      // A[0]: Timestamp | B[1]: Mês | C[2]: Ano | D[3]: QR Code | E[4]: Chave Pix | F[5]: Valor | G[6]: Suplementar | H[7]: Vencimento
      
      const mesRef = linha[1];
      const anoRef = linha[2];
      const linkQrOriginal = linha[3]; 
      const chavePix = linha[4]; 
      const valorBruto = linha[5];
      const inputSuplementar = linha[6]; 
      const dataVencimento = linha[7];   

      if (mesRef || valorBruto) {
        let valorNumerico = tratarValor(valorBruto); 

        // 1. MÊS COMPOSTO
        let textoMesComposto = mesRef;
        let indiceSup = 0;
        if (inputSuplementar && inputSuplementar.toString().trim() !== "") {
            indiceSup = parseInt(inputSuplementar.toString().trim());
        }
        if (indiceSup > 0) {
            textoMesComposto = `${mesRef} (${indiceSup})`;
        }

        // 2. LINK QR CODE
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

        let textoVencimento = dataVencimento || "-";

        // 3. CRIAÇÃO DAS CÉLULAS
        let rtMes = SpreadsheetApp.newRichTextValue().setText(textoMesComposto).setTextStyle(estiloNormal).build();
        let rtAno = SpreadsheetApp.newRichTextValue().setText(anoRef).setTextStyle(estiloNormal).build();
        let rtVencimento = SpreadsheetApp.newRichTextValue().setText(textoVencimento).setTextStyle(estiloNormal).build();
        let rtChave = SpreadsheetApp.newRichTextValue().setText(chavePix || "-").setTextStyle(estiloNormal).build();
        
        let textoValor = valorNumerico.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
        let rtValor = SpreadsheetApp.newRichTextValue().setText(textoValor).setTextStyle(estiloNormal).build(); 

        // 4. MONTAGEM DA LINHA FINAL
        saidaRichText.push([rtMes, rtAno, rtVencimento, rtValor, rtQr, rtChave]);
        valoresNumericos.push([valorNumerico]);
      }
    });
  }

  // --- 5. ESCREVER ---
  if (saidaRichText.length > 0) {
    const numLinhas = saidaRichText.length;
    
    // Escreve as 6 colunas
    abaDestino.getRange(1, 1, numLinhas, 6).setRichTextValues(saidaRichText);
    
    // Formata Valor (Coluna D = 4)
    abaDestino.getRange(1, 4, numLinhas, 1).setNumberFormat("R$ #,##0.00");
    
    // Sobrescreve números
    if(valoresNumericos.length > 0){
       abaDestino.getRange(1, 4, valoresNumericos.length, 1).setValues(valoresNumericos);
    }
    
    // Alinhamentos
    abaDestino.getRange(1, 1, numLinhas, 6).setHorizontalAlignment("center");
    abaDestino.getRange(1, 6, numLinhas, 1).setHorizontalAlignment("left");
    
    abaDestino.autoResizeColumns(1, 6);
  }
}