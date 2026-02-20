function processarEntradaFluxoCaixa(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomeAba = "📥 DB_Pagamentos"; 
  let aba = ss.getSheetByName(nomeAba);

  // 1. Se a aba não existir, cria
  if (!aba) {
    aba = ss.insertSheet(nomeAba);
  }

  // 2. GARANTIA DE CABEÇALHO
  aba.getRange(1, 1, 1, 8).setValues([[
    "Mês Ref", 
    "Ano", 
    "Marco (Valor)", "Data Pag.", 
    "Janaína (Valor)", "Data Pag.", 
    "Adriana (Valor)", "Data Pag."
  ]]);
  aba.setFrozenRows(1);
  aba.getRange("A1:H1").setFontWeight("bold").setBackground("#000000").setFontColor("#ffffff");

  // 3. Busca se já existe uma linha de dados para esse Mês/Ano
  // Nota: Lemos os dados antes de gravar para encontrar a posição
  let dadosSheet = aba.getDataRange().getValues();
  let linhaEncontrada = -1;

  for (let i = 1; i < dadosSheet.length; i++) {
    let mesLinha = String(dadosSheet[i][0]).trim(); 
    let anoLinha = String(dadosSheet[i][1]).trim();
    
    if (mesLinha === String(dados.mesRef).trim() && anoLinha === String(dados.anoRef).trim()) {
      linhaEncontrada = i + 1; 
      break;
    }
  }

  // 4. Se não achou, cria nova linha (com 0 e Emoji)
  if (linhaEncontrada === -1) {
    aba.appendRow([
      dados.mesRef, 
      dados.anoRef, 
      0, "😢",  
      0, "😢",  
      0, "😢"   
    ]);
    linhaEncontrada = aba.getLastRow();
  }

  // 5. Mapa das Colunas
  const mapaColunas = {
    "Marco":   { colValor: 3, colData: 4 },
    "Janaina": { colValor: 5, colData: 6 },
    "Adriana": { colValor: 7, colData: 8 }
  };

  // 6. Grava os dados recebidos
  dados.pagamentos.forEach(pgto => {
    let nomeChave = pgto.pessoa.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (nomeChave.toLowerCase().includes("marco")) nomeChave = "Marco";
    else if (nomeChave.toLowerCase().includes("janaina")) nomeChave = "Janaina";
    else if (nomeChave.toLowerCase().includes("adriana")) nomeChave = "Adriana";

    const coords = mapaColunas[nomeChave];
    
    if (coords) {
      let dataParts = pgto.data.split("-");
      let dataObj = new Date(dataParts[0], dataParts[1] - 1, dataParts[2]);
      let valorLimpo = parseFloat(String(pgto.valor).replace(/\./g, "").replace(",", "."));

      aba.getRange(linhaEncontrada, coords.colValor).setValue(valorLimpo);
      aba.getRange(linhaEncontrada, coords.colData).setValue(dataObj);
    }
  });

  // 7. Formatação Visual (na linha editada)
  aba.getRange(linhaEncontrada, 3).setNumberFormat("R$ #,##0.00");
  aba.getRange(linhaEncontrada, 5).setNumberFormat("R$ #,##0.00");
  aba.getRange(linhaEncontrada, 7).setNumberFormat("R$ #,##0.00");

  let padraoData = "dd/MM";
  aba.getRange(linhaEncontrada, 4).setNumberFormat(padraoData);
  aba.getRange(linhaEncontrada, 6).setNumberFormat(padraoData);
  aba.getRange(linhaEncontrada, 8).setNumberFormat(padraoData);
  
  aba.getRange(linhaEncontrada, 1, 1, 8).setHorizontalAlignment("center");

  // =================================================================
  // 8. ORDENAÇÃO AUTOMÁTICA (O Passo Novo)
  // =================================================================
  // Pega todos os dados (exceto cabeçalho) e reordena na memória
  const ultimaLinha = aba.getLastRow();
  
  if (ultimaLinha > 1) {
    const rangeDados = aba.getRange(2, 1, ultimaLinha - 1, 8);
    const valores = rangeDados.getValues();

    valores.sort((a, b) => {
      // a[0] = Mês (ex: "Janeiro" ou "Janeiro (1)")
      // a[1] = Ano (ex: 2026)
      // A função ordenarCobrancasPorPeriodo já sabe lidar com o "(1)"
      return ordenarCobrancasPorPeriodo(a[0], a[1]) - ordenarCobrancasPorPeriodo(b[0], b[1]);
    });

    // Devolve os dados ordenados para a planilha
    rangeDados.setValues(valores);
  }
}