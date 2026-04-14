/**
 * Gera o 3º Dashboard: Resumo do Fluxo de Caixa (Quem pagou quem)
 * Localização: Colunas M, N, O, P, Q
 */
function gerarDashboardFluxoCaixa() {

  // --- CONFIGURAÇÕES ---
  const nomeAbaDados = "💰 Fluxo_Caixa_Dados_Brutos";
  const nomeAbaRelatorio = "⭐ Dashboard Gestão";
  
  // Coluna M é a 13ª letra do alfabeto
  const colunaInicial = 13; 

  const cabecalho = [
    "Mês/Ano Referência",
    "Cobrança",
    "Marco",
    "Janaína",
    "Adriana"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDados = ss.getSheetByName(nomeAbaDados);
  let abaRelatorio = ss.getSheetByName(nomeAbaRelatorio);

  if (!abaRelatorio) abaRelatorio = ss.insertSheet(nomeAbaRelatorio);

  // --- 1. LIMPAR ÁREA (Da coluna 13 até 17, linhas infinitas) ---
  abaRelatorio.getRange(1, colunaInicial, 200, 5).clear();

  // --- 2. TÍTULO ---
  abaRelatorio.getRange(1, colunaInicial, 1, 5) 
    .merge()
    .setValue("💰 Resumo Fluxo de Caixa")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground("white")
    .setFontColor("black")
    .setHorizontalAlignment("Left");

  // --- 3. CABEÇALHO ---
  abaRelatorio.getRange(2, colunaInicial, 1, 5) 
    .setValues([cabecalho])
    .setFontWeight("bold")
    .setBackground("#000000") 
    .setFontColor("#FFFFFF"); 

  if (!abaDados) return;

  // --- 4. LER DADOS ---
  const dados = abaDados.getDataRange().getValues();
  if (dados.length < 2) return;

  const linhas = dados.slice(1); // Pula cabeçalho original
  const listaProcessada = [];

  // --- 5. PROCESSAR DADOS ---
  linhas.forEach(linha => {
    // Mapeamento baseado no seu layout:
    // A=Mes, B=Ano, C=Marco(V), D=Marco(D), E=Janaina(V), F=Janaina(D), G=Adriana(V), H=Adriana(D)
    
    let mesCru = String(linha[0]); 
    const ano = linha[1];
    
    // Dados Marco
    const valorMarco = linha[2];
    const dataMarco = linha[3];
    
    // Dados Janaina
    const valorJanaina = linha[4];
    const dataJanaina = linha[5];
    
    // Dados Adriana
    const valorAdriana = linha[6];
    const dataAdriana = linha[7];

    if (!mesCru || !ano) return;

    // Lógica de "Mês (Extra)"
    let mesLimpo = mesCru.trim();
    let indiceNum = 0;
    const match = mesCru.match(/\((\d+)\)/);
    if (match) {
      indiceNum = parseInt(match[1]); 
      mesLimpo = mesCru.replace(/\s*\(\d+\)/, "").trim(); 
    }
    mesLimpo = mesLimpo.toLowerCase();

    // Função auxiliar para determinar status visual
    const checkStatus = (valor, data) => {
      // Se a data contiver o emoji triste ou for vazia, ou valor for 0
      if (String(data).includes("😢") || !data || valor === 0 || valor === "") {
        return "❌";
      }
      return "✅";
    };

    listaProcessada.push({
      mes: mesLimpo,
      ano: ano,
      indice: indiceNum,
      // Status Calculados
      statusMarco: checkStatus(valorMarco, dataMarco),
      statusJanaina: checkStatus(valorJanaina, dataJanaina),
      statusAdriana: checkStatus(valorAdriana, dataAdriana)
    });
  });

  if (listaProcessada.length === 0) return;

  // --- 6. ORDENAR ---
  listaProcessada.sort((a, b) => {
    return (
      ordenarCobrancasPorPeriodo(a.mes, a.ano, a.indice) -
      ordenarCobrancasPorPeriodo(b.mes, b.ano, b.indice)
    );
  });

  // --- 7. PREPARAR SAÍDA ---
  const saida = [];
  
  listaProcessada.forEach(item => {
    const nomeMes = item.mes.charAt(0).toUpperCase() + item.mes.slice(1);
    const mesAno = `${nomeMes} ${item.ano}`;
    
    let tipo = item.indice === 0 ? "Padrão" : `Extra ${item.indice}`;

    saida.push([
      mesAno,
      tipo,
      item.statusMarco,
      item.statusJanaina,
      item.statusAdriana
    ]);
  });

  // --- 8. ESCREVER DADOS ---
  const rangeTabela = abaRelatorio.getRange(3, colunaInicial, saida.length, 5);
  rangeTabela.setValues(saida);

  // Formatação Alinhamento
  rangeTabela.setHorizontalAlignment("center"); // Centraliza os Checks
  abaRelatorio.getRange(3, colunaInicial, saida.length, 2).setHorizontalAlignment("left"); // Alinha textos à esquerda

  // Ajuste de largura
  abaRelatorio.autoResizeColumns(colunaInicial, 5);
}