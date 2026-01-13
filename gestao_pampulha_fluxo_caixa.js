function gestao_pampulha_fluxo_caixa_dados_brutos() {
  // --- CONFIGURAÇÕES ---
  const idPlanilhaOrigem = "1p9gJn3NRDV48y06QVc97YlUsxZCLc4W001qO20daVeQ"; 
  const nomeAbaOrigem = "Respostas ao formulário 1"; 
  const nomeAbaDestino = "💰 Fluxo_Caixa_Dados_Brutos"; 

  // --- ACESSAR PLANILHAS ---
  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  let abaDestino = ssDestino.getSheetByName(nomeAbaDestino);

  if (!abaDestino) {
    abaDestino = ssDestino.insertSheet(nomeAbaDestino);
  }

  // --- ESTILOS ---
  const estiloNormal = SpreadsheetApp.newTextStyle()
    .setUnderline(false)
    .setForegroundColor("black")
    .build();

  const estiloLink = SpreadsheetApp.newTextStyle()
    .setUnderline(true)
    .setForegroundColor("#1155cc")
    .build();

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

  function processarPessoa(pagou, valorStr, dataStr, linkStr) {
    let rtValor, rtData;
    const pagouNormalizado = pagou ? pagou.toString().trim().toLowerCase() : "";

    if (pagouNormalizado === "sim") {
      let valNum = tratarValor(valorStr);
      let textoVal = valNum.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

      let builderValor = SpreadsheetApp.newRichTextValue().setText(textoVal);
      
      if (linkStr && linkStr.toString().includes("http")) {
        builderValor.setLinkUrl(linkStr);
        builderValor.setTextStyle(estiloLink);
      } else {
        builderValor.setLinkUrl(null);
        builderValor.setTextStyle(estiloNormal);
      }
      rtValor = builderValor.build();

      let textoData = (dataStr && dataStr.toString().trim() !== "") ? dataStr : "-";
      rtData = SpreadsheetApp.newRichTextValue().setText(textoData).setTextStyle(estiloNormal).build();

    } else {
      let zeroVal = (0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
      rtValor = SpreadsheetApp.newRichTextValue().setText(zeroVal).setTextStyle(estiloNormal).build();
      rtData = SpreadsheetApp.newRichTextValue().setText("-").setTextStyle(estiloNormal).build();
    }

    return { rtValor, rtData, valorNumerico: (pagouNormalizado === "sim" ? tratarValor(valorStr) : 0) };
  }

  // --- LEITURA E PROCESSAMENTO ---
  try {
    // 1. LIMPEZA IMEDIATA: Apaga tudo antes de começar
    abaDestino.clear();

    const ssOrigem = SpreadsheetApp.openById(idPlanilhaOrigem);
    const abaOrigem = ssOrigem.getSheetByName(nomeAbaOrigem);
    const dadosOrigem = abaOrigem.getDataRange().getDisplayValues();

    // 2. PREPARA O CABEÇALHO (Garante que a tabela tenha títulos mesmo vazia)
    const saidaRichText = [];
    const valoresNumericos = []; // Para formatar moeda depois

    const cabecalho = [
      "Mês Ref", "Ano", 
      "Marco (Valor)", "Data Pag.", 
      "Janaína (Valor)", "Data Pag.", 
      "Adriana (Valor)", "Data Pag."
    ].map(txt => SpreadsheetApp.newRichTextValue().setText(txt).setTextStyle(estiloNormal).build());
    
    saidaRichText.push(cabecalho);
    valoresNumericos.push([null, null, null, null, null, null, null, null]); 

    // 3. PROCESSA DADOS (Se houver linhas além do cabeçalho original)
    if (dadosOrigem.length >= 2) {
      
      let linhasDados = dadosOrigem.slice(1);

      // Ordenação
      linhasDados.sort(function(a, b) {
        const anoA = parseInt(a[2]) || 0;
        const anoB = parseInt(b[2]) || 0;
        const mesA = obterNumeroMes(a[1]);
        const mesB = obterNumeroMes(b[1]);

        if (anoA !== anoB) return anoA - anoB; 
        return mesA - mesB;   
      });

      // Loop pelos dados
      linhasDados.forEach(linha => {
        const mesRef = linha[1];
        const anoRef = linha[2];

        // Se a linha não tiver Mês nem Ano, ignora (evita sujeira)
        if (mesRef || anoRef) {
          const marco = processarPessoa(linha[3], linha[4], linha[5], linha[6]);
          const janaina = processarPessoa(linha[7], linha[8], linha[9], linha[10]);
          const adriana = processarPessoa(linha[11], linha[12], linha[13], linha[14]);

          const rtMes = SpreadsheetApp.newRichTextValue().setText(mesRef).setTextStyle(estiloNormal).build();
          const rtAno = SpreadsheetApp.newRichTextValue().setText(anoRef).setTextStyle(estiloNormal).build();

          saidaRichText.push([
            rtMes, rtAno,
            marco.rtValor, marco.rtData,
            janaina.rtValor, janaina.rtData,
            adriana.rtValor, adriana.rtData
          ]);

          valoresNumericos.push([
            null, null, 
            marco.valorNumerico, null, 
            janaina.valorNumerico, null, 
            adriana.valorNumerico, null
          ]);
        }
      });
    }

    // --- 4. ESCREVER NA ABA DESTINO ---
    if (saidaRichText.length > 0) {
      const numLinhas = saidaRichText.length;
      const numCols = saidaRichText[0].length;
      
      // Escreve RichTexts (Texto + Links)
      abaDestino.getRange(1, 1, numLinhas, numCols).setRichTextValues(saidaRichText);

      // Se tiver dados numéricos (mais que 1 linha), aplica formatação
      if (numLinhas > 1) {
        const rangeMarco = abaDestino.getRange(2, 3, numLinhas - 1, 1);
        const rangeJanaina = abaDestino.getRange(2, 5, numLinhas - 1, 1);
        const rangeAdriana = abaDestino.getRange(2, 7, numLinhas - 1, 1);
        
        const formatoMoeda = "R$ #,##0.00";
        rangeMarco.setNumberFormat(formatoMoeda);
        rangeJanaina.setNumberFormat(formatoMoeda);
        rangeAdriana.setNumberFormat(formatoMoeda);
      }

      // Ajuste de largura
      abaDestino.autoResizeColumns(1, 8);
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro: " + e.message);
  }
}