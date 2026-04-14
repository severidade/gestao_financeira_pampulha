function gestao_pampulha_fluxo_caixa_dados_brutos() {
  // --- CONFIGURAÇÕES ---
  const nomeAbaOrigem = "📥 DB_Pagamentos"; // A nova aba criada pelo painel
  const nomeAbaDestino = "💰 Fluxo_Caixa_Dados_Brutos"; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaOrigem = ss.getSheetByName(nomeAbaOrigem);
  let abaDestino = ss.getSheetByName(nomeAbaDestino);

  if (!abaOrigem) {
    SpreadsheetApp.getUi().alert("Aba de Banco de Dados de Pagamentos não encontrada. Use o Painel para criar.");
    return;
  }
  if (!abaDestino) abaDestino = ss.insertSheet(nomeAbaDestino);

  // --- 1. LIMPAR DESTINO ---
  abaDestino.clear();

  // --- 2. LER DADOS DO BANCO ---
  const dadosBrutos = abaOrigem.getDataRange().getValues();
  if (dadosBrutos.length < 2) return; // Só tem cabeçalho

  // --- 3. ESTRUTURA DE PIVOT (TRANSFORMAR LINHAS EM COLUNAS) ---
  // Chave = "Janeiro (1)|2026"
  // Valor = { Marco: {val, data}, Janaina: {val, data}, ... }
  let agrupamento = {};

  // Pula cabeçalho (i=1)
  for (let i = 1; i < dadosBrutos.length; i++) {
    let mes = dadosBrutos[i][1]; // Coluna B
    let ano = dadosBrutos[i][2]; // Coluna C
    let pessoa = dadosBrutos[i][3]; // Coluna D
    let dataPag = dadosBrutos[i][4]; // Coluna E
    let valor = dadosBrutos[i][5]; // Coluna F

    if (!mes || !ano) continue;

    let chave = `${mes}|${ano}`;

    if (!agrupamento[chave]) {
      agrupamento[chave] = {
        mes: mes,
        ano: ano,
        Marco: null,
        Janaina: null, // Atenção ao acento no DB vs Código
        Adriana: null
      };
    }

    // Normaliza nome para evitar erro de digitação "Janaína" vs "Janaina"
    let pessoaNorm = String(pessoa).normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // Remove acentos
    if (pessoaNorm.toLowerCase().includes("marco")) pessoaNorm = "Marco";
    if (pessoaNorm.toLowerCase().includes("janaina")) pessoaNorm = "Janaina";
    if (pessoaNorm.toLowerCase().includes("adriana")) pessoaNorm = "Adriana";

    agrupamento[chave][pessoaNorm] = {
      valor: valor,
      data: dataPag
    };
  }

  // --- 4. PREPARAR SAÍDA ---
  let listaSaida = Object.values(agrupamento);
  
  // Ordenar usando sua função
  listaSaida.sort((a, b) => {
    // Extrai o índice do texto "Janeiro (1)" se houver
    let indA = 0, indB = 0;
    let mesCleanA = a.mes, mesCleanB = b.mes;
    
    let matchA = a.mes.match(/\((\d+)\)/);
    if(matchA) { indA = matchA[1]; mesCleanA = a.mes.split("(")[0].trim(); }

    let matchB = b.mes.match(/\((\d+)\)/);
    if(matchB) { indB = matchB[1]; mesCleanB = b.mes.split("(")[0].trim(); }

    return ordenarCobrancasPorPeriodo(mesCleanA, a.ano, indA) - ordenarCobrancasPorPeriodo(mesCleanB, b.ano, indB);
  });

  // Montar array para planilha
  // Cabeçalho
  let linhasFinais = [];
  linhasFinais.push(["Mês Ref", "Ano", "Marco (Valor)", "Data Pag.", "Janaína (Valor)", "Data Pag.", "Adriana (Valor)", "Data Pag."]);

  let formatos = []; // Guardar onde aplicar formato moeda/data

  listaSaida.forEach(item => {
    let linha = [
       item.mes,
       item.ano,
       item.Marco ? item.Marco.valor : 0,
       item.Marco ? item.Marco.data : "-",
       item.Janaina ? item.Janaina.valor : 0,
       item.Janaina ? item.Janaina.data : "-",
       item.Adriana ? item.Adriana.valor : 0,
       item.Adriana ? item.Adriana.data : "-"
    ];
    linhasFinais.push(linha);
  });

  // --- 5. ESCREVER ---
  if (linhasFinais.length > 0) {
    let numL = linhasFinais.length;
    let numC = linhasFinais[0].length;
    
    abaDestino.getRange(1, 1, numL, numC).setValues(linhasFinais);
    
    // Formatação
    abaDestino.getRange(1, 1, 1, numC).setFontWeight("bold").setBackground("#eee");
    
    // Colunas de Valor (C, E, G -> 3, 5, 7)
    if (numL > 1) {
      abaDestino.getRange(2, 3, numL - 1, 1).setNumberFormat("R$ #,##0.00");
      abaDestino.getRange(2, 5, numL - 1, 1).setNumberFormat("R$ #,##0.00");
      abaDestino.getRange(2, 7, numL - 1, 1).setNumberFormat("R$ #,##0.00");
      
      // Colunas de Data (D, F, H -> 4, 6, 8)
      let padraoData = "dd/MM";
      abaDestino.getRange(2, 4, numL - 1, 1).setNumberFormat(padraoData);
      abaDestino.getRange(2, 6, numL - 1, 1).setNumberFormat(padraoData);
      abaDestino.getRange(2, 8, numL - 1, 1).setNumberFormat(padraoData);
    }
    
    abaDestino.autoResizeColumns(1, numC);
  }
}