/**
 * Gera o Dashboard de Relatório de Acertos (Direita)
 * - Aceita QUALQUER texto na coluna de status como "Enviado/Pago"
 */
function gerarDashboardRelatorioAcertos() {

  // --- CONFIGURAÇÕES ---
  const nomeAbaDados = "🤝 Acertos_Mensais_Dados_Brutos";
  const nomeAbaRelatorio = "⭐ Dashboard Gestão";
  const linhaInicial = 2; 
  const colunaInicial = 7; // Coluna G

  const cabecalho = [
    "Mês/Ano Referência",
    "Cobrança",
    "Vencimento",
    "Valor Cobrado",
    "Enviado?" 
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDados = ss.getSheetByName(nomeAbaDados);
  let abaRelatorio = ss.getSheetByName(nomeAbaRelatorio);

  if (!abaRelatorio) abaRelatorio = ss.insertSheet(nomeAbaRelatorio);

  // --- 1. LIMPAR TUDO ---
  abaRelatorio.getRange(1, colunaInicial, 200, 5).clear();

  // --- 2. TÍTULO ---
  abaRelatorio.getRange(1, colunaInicial, 1, 5) 
    .merge()
    .setValue("🤝 Resumo de Acertos")
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
  // const dados = abaDados.getDataRange().getDisplayValues();
  const dados = abaDados.getDataRange().getValues();
  if (dados.length < 2) return;

  const linhas = dados.slice(1);
  const listaProcessada = [];

  // --- 5. PROCESSAR DADOS ---
  linhas.forEach(linha => {
    let mesCru = String(linha[0]); 
    const ano = linha[1];
    const vencimento = linha[2]; 
    const valorStr = linha[3]; 
    const statusEnvio = linha[6]; // LER COLUNA G

    if (!mesCru || !ano) return;

    let mesLimpo = mesCru.trim();
    let indiceNum = 0;
    const match = mesCru.match(/\((\d+)\)/);
    if (match) {
      indiceNum = parseInt(match[1]); 
      mesLimpo = mesCru.replace(/\s*\(\d+\)/, "").trim(); 
    }
    mesLimpo = mesLimpo.toLowerCase();

    // let valor = parseFloat(String(valorStr).replace("R$", "").replace(/\./g, "").replace(",", ".").trim()) || 0;
    let valor = Number(linha[3]) || 0;

    listaProcessada.push({
      mes: mesLimpo,
      ano: ano,
      indice: indiceNum,
      vencimento: vencimento,
      valor: valor,
      status: statusEnvio
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
  const matrizFundos = []; 

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0); 

listaProcessada.forEach(item => {
  const nomeMes = item.mes.charAt(0).toUpperCase() + item.mes.slice(1);
  const mesAno = `${nomeMes} ${item.ano}`;
  
  let tipo = item.indice === 0 ? "Padrão" : `Extra ${item.indice}`;
  
  let corFundo = "white"; 

  // 🔥 TRATAMENTO CORRETO DE DATA
  let dataConta = null;

  if (item.vencimento instanceof Date) {
    dataConta = item.vencimento;

  } else if (
    typeof item.vencimento === "string" &&
    item.vencimento.includes("/")
  ) {
    const partes = item.vencimento.split("/");
    dataConta = new Date(partes[2], partes[1] - 1, partes[0]);
  }

  if (dataConta && dataConta < hoje) {
    corFundo = "#EEEEEE";
  }

  // --- LÓGICA VISUAL DO STATUS ---
  let visualStatus = "-";
  
  if (
    item.status &&
    String(item.status).trim() !== "" &&
    String(item.status).trim() !== "-"
  ) {
    visualStatus = "✅";
  }

  saida.push([mesAno, tipo, item.vencimento, item.valor, visualStatus]);
  matrizFundos.push([corFundo, corFundo, corFundo, corFundo, corFundo]);
});

  // --- 8. ESCREVER DADOS ---
  const rangeTabela = abaRelatorio.getRange(3, colunaInicial, saida.length, 5);
  
  rangeTabela.setValues(saida);
  rangeTabela.setHorizontalAlignment("left");
  abaRelatorio.getRange(3, colunaInicial + 4, saida.length, 1).setHorizontalAlignment("center");
  rangeTabela.setBackgrounds(matrizFundos); 

  abaRelatorio.getRange(3, colunaInicial + 3, saida.length, 1).setNumberFormat("R$ #,##0.00");
  abaRelatorio.autoResizeColumns(colunaInicial, 5);

  // --- 9. RODAPÉ ---
  const linhaRodape = 3 + saida.length; 
  const celulaRodape = abaRelatorio.getRange(linhaRodape, colunaInicial, 1, 5);
  celulaRodape
    .merge() 
    .setValue("Linhas em cinza indicam que a data de vencimento já passou, mas não confirmam o pagamento. Ir para a aba 💰 Fluxo_Caixa_Dados_Brutos e cadastrar o fechamento.")
    .setFontSize(8).setFontStyle("italic").setFontColor("#333") 
    .setBackground("white").setHorizontalAlignment("left").setVerticalAlignment("middle").setWrap(true);          
}