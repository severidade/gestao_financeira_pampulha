function gerar_relatorio_passivos() {
  // --- CONFIGURAÇÕES ---
  const nomeAbaDados = "💸 Passivos_Dados_Brutos";
  const nomeAbaRelatorio = "⭐ Dashboard Gestão"; // MUDANÇA: Aponta para o Dashboard
  
  // CONFIGURAÇÃO DE POSIÇÃO NO DASHBOARD
  const linhaInicial = 1; // A tabela começará na linha 15 (para não apagar o topo)
  const colunaInicial = 1; // A tabela começará na coluna A (1)
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDados = ss.getSheetByName(nomeAbaDados);
  let abaRelatorio = ss.getSheetByName(nomeAbaRelatorio);

  if (!abaRelatorio) {
    abaRelatorio = ss.insertSheet(nomeAbaRelatorio);
  }

  // --- 1. LER DADOS BRUTOS ---
  const dados = abaDados.getDataRange().getDisplayValues();
  
  if (dados.length < 2) return;

  const linhas = dados.slice(1); 

  // --- 2. AGRUPAR E SOMAR ---
  const agrupamento = {};

  linhas.forEach(linha => {
    // A=Mês(0), B=Ano(1), C=Serviço, D=Valor(3), E=Recibo, F=Suplementar(5)
    
    let mes = linha[0];
    let ano = linha[1];
    let valorStr = linha[3];
    let indiceSuplementar = linha[5]; // Coluna F

    if (!mes || !ano) return;
    if (!indiceSuplementar) indiceSuplementar = "0"; 

    let valor = parseFloat(valorStr.replace("R$", "").replace(/\./g, "").replace(",", ".").trim());
    if (isNaN(valor)) valor = 0;

    let chave = `${mes}|${ano}|${indiceSuplementar}`;

    if (!agrupamento[chave]) {
      agrupamento[chave] = {
        mes: mes,
        ano: ano,
        indice: parseInt(indiceSuplementar),
        total: 0
      };
    }
    agrupamento[chave].total += valor;
  });

  // --- 3. TRANSFORMAR EM LISTA E ORDENAR ---
  let listaFinal = Object.values(agrupamento);

  const mapaMeses = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "marco": 3, "abril": 4, "maio": 5, "junho": 6,
    "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
  };

  listaFinal.sort((a, b) => {
    // 1. Ano
    if (parseInt(a.ano) !== parseInt(b.ano)) return parseInt(a.ano) - parseInt(b.ano);
    // 2. Mês
    let mesNumA = mapaMeses[a.mes.toLowerCase()] || 0;
    let mesNumB = mapaMeses[b.mes.toLowerCase()] || 0;
    if (mesNumA !== mesNumB) return mesNumA - mesNumB;
    // 3. Suplementar
    return a.indice - b.indice;
  });

  // --- 4. PREPARAR DADOS PARA EXIBIÇÃO ---
  const saida = [];
  saida.push(["Mês/Ano Referência", "Cobrança", "Total da Casa", "Valor Individual (1/3)"]);

  listaFinal.forEach(item => {
    let colunaMesAno = `${item.mes} ${item.ano}`;
    let colunaTipo = item.indice === 0 ? "padrao" : `suplementar ${item.indice}`;
    let total = item.total;
    let individual = item.total / 3;

    saida.push([colunaMesAno, colunaTipo, total, individual]);
  });

  // --- 5. ESCREVER NO DASHBOARD (ATUALIZADO) ---
  
  // Limpeza de Segurança: Limpa apenas a área onde a tabela antiga estava
  // (Do início definido até 100 linhas para baixo, nas 4 colunas da tabela)
  // Isso evita que sobrem dados antigos se a nova tabela for menor
  abaRelatorio.getRange(linhaInicial, colunaInicial, 100, 4).clear(); 
  
  if (saida.length > 0) {
    // Escreve os novos dados na posição definida
    abaRelatorio.getRange(linhaInicial, colunaInicial, saida.length, 4).setValues(saida);
    
    // Formatação Visual
    // Cabeçalho
    abaRelatorio.getRange(linhaInicial, colunaInicial, 1, 4)
      // .setFontWeight("bold")
      // .setBackground("#EFEFEF")
      .setBorder(true, true, true, true, null, null); 
    
    // === NOVO: ALINHAMENTO À ESQUERDA ===
    rangeTabela.setHorizontalAlignment("left");
    
    // Colunas de Valor (Começa na linhaInicial + 1)
    if (saida.length > 1) {
      abaRelatorio.getRange(linhaInicial + 1, colunaInicial + 2, saida.length - 1, 2).setNumberFormat("R$ #,##0.00");
    }
    
    // Ajuste de largura (Opcional - pode remover se estragar o layout do dashboard)
    // abaRelatorio.autoResizeColumns(colunaInicial, 4);
  }
}