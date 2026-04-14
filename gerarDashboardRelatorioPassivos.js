function gerarDashboardRelatorioPassivos() {

  // --- CONFIGURAÇÕES ---
  const nomeAbaDados = "💸 Passivos_Dados_Brutos";
  const nomeAbaRelatorio = "⭐ Dashboard Gestão";

  // A tabela começa visualmente na linha 1 (Título), Cabeçalho na 2
  const colunaInicial = 1; // Coluna A

  const cabecalho = [
    "Mês/Ano Referência",
    "Cobrança",
    "Total da Casa",
    "Valor Individual (1/3)"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDados = ss.getSheetByName(nomeAbaDados);
  let abaRelatorio = ss.getSheetByName(nomeAbaRelatorio);

  if (!abaRelatorio) {
    abaRelatorio = ss.insertSheet(nomeAbaRelatorio);
  }

  // --- 1. LIMPAR TUDO (DA LINHA 1 PARA BAIXO) ---
  abaRelatorio
    .getRange(1, colunaInicial, 100, 4) 
    .clear();

  // --- 2. ESCREVER TÍTULO (LINHA 1 - MESCLADA) ---
  abaRelatorio.getRange(1, colunaInicial, 1, 4) // Seleciona A1 até D1
    .merge() // <--- Mescla as células
    .setValue("💸 Resumo de Despesas e Rateio")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground("white")
    .setHorizontalAlignment("left") // Centraliza o texto na área mesclada
    .setVerticalAlignment("middle");

  // --- 3. ESCREVER CABEÇALHO (LINHA 2 - PRETO E BRANCO) ---
  abaRelatorio
    .getRange(2, colunaInicial, 1, 4)
    .setValues([cabecalho])
    .setFontWeight("bold")
    .setBackground("#000000") // Fundo Preto
    .setFontColor("#FFFFFF"); // Texto Branco

  if (!abaDados) { abaRelatorio.autoResizeColumns(colunaInicial, 4); return; }

  // --- 4. LER DADOS ---
  const dados = abaDados.getDataRange().getDisplayValues();
  if (dados.length < 2) { abaRelatorio.autoResizeColumns(colunaInicial, 4); return; }

  const linhas = dados.slice(1);
  const agrupamento = {};

  // --- 5. AGRUPAR E SOMAR ---
  linhas.forEach(linha => {
    let mesCru = String(linha[0]); 
    const ano = linha[1];
    const valorStr = linha[3];
    
    // Lógica de Detecção e Limpeza
    let mesLimpo = mesCru.trim();
    let indiceNum = 0;

    const match = mesCru.match(/\((\d+)\)/);
    if (match) {
      indiceNum = parseInt(match[1]); 
      mesLimpo = mesCru.replace(/\s*\(\d+\)/, "").trim(); 
    }

    mesLimpo = mesLimpo.toLowerCase(); 

    if (!mesLimpo || !ano) return;

    let valor = parseFloat(
      String(valorStr)
        .replace("R$", "")
        .replace(/\./g, "")
        .replace(",", ".")
        .trim()
    );
    if (isNaN(valor)) valor = 0;

    const chave = `${mesLimpo}|${ano}|${indiceNum}`;

    if (!agrupamento[chave]) {
      agrupamento[chave] = {
        mes: mesLimpo,
        ano: ano,
        indice: indiceNum,
        total: 0
      };
    }

    agrupamento[chave].total += valor;
  });

  const listaFinal = Object.values(agrupamento);

  if (listaFinal.length === 0) {
    abaRelatorio.autoResizeColumns(colunaInicial, 4);
    return;
  }

  // --- 6. ORDENAR (CRESCENTE) ---
  listaFinal.sort((a, b) => {
    return (
      ordenarCobrancasPorPeriodo(a.mes, a.ano, a.indice) -
      ordenarCobrancasPorPeriodo(b.mes, b.ano, b.indice)
    );
  });

  // --- 7. PREPARAR SAÍDA ---
  const saida = [];

  listaFinal.forEach(item => {
    const nomeMes = item.mes.charAt(0).toUpperCase() + item.mes.slice(1);
    const mesAno = `${nomeMes} ${item.ano}`;

    let tipo = "";
    if (item.indice === 0) {
        tipo = "Padrão";
    } else {
        tipo = `Extra ${item.indice}`;
    }

    const total = item.total;
    const individual = total / 3;

    saida.push([mesAno, tipo, total, individual]);
  });

  // --- 8. ESCREVER DADOS (A PARTIR DA LINHA 3) ---
  const rangeTabela = abaRelatorio.getRange(
    3, // Linha 3
    colunaInicial,
    saida.length,
    4
  );

  rangeTabela.setValues(saida);
  rangeTabela.setHorizontalAlignment("left");

  abaRelatorio.getRange(3, colunaInicial + 2, saida.length, 2)
    .setNumberFormat("R$ #,##0.00");

  abaRelatorio.autoResizeColumns(colunaInicial, 4);
}