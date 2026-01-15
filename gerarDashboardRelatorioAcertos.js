/**
 * Gera o Dashboard de Relatório de Acertos (Direita)
 * - Cabeçalho: Preto com texto Branco
 * - Contas vencidas: Fundo Cinza Claro (#EEEEEE)
 * - Rodapé: Explicação sobre o status de vencimento
 */
function gerarDashboardRelatorioAcertos() {

  // --- CONFIGURAÇÕES ---
  const nomeAbaDados = "🤝 Acertos_Mensais_Dados_Brutos";
  const nomeAbaRelatorio = "⭐ Dashboard Gestão";
  const linhaInicial = 2; // Onde fica o cabeçalho
  const colunaInicial = 7; // Coluna G

  const cabecalho = [
    "Mês/Ano Referência",
    "Cobrança",
    "Vencimento",
    "Valor Cobrado"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDados = ss.getSheetByName(nomeAbaDados);
  let abaRelatorio = ss.getSheetByName(nomeAbaRelatorio);

  if (!abaRelatorio) abaRelatorio = ss.insertSheet(nomeAbaRelatorio);

  // --- 1. LIMPAR TUDO (Margem de segurança maior para o rodapé) ---
  abaRelatorio.getRange(1, colunaInicial, 200, 4).clear();

  // --- 2. TÍTULO (LINHA 1) ---
  abaRelatorio.getRange(1, colunaInicial)
    .setValue("🤝 Resumo de Acertos")
    .setFontWeight("bold").setFontSize(14).setBackground("white").setFontColor("black");

  // --- 3. CABEÇALHO DA TABELA (LINHA 2) - ESTILO NOVO ---
  abaRelatorio.getRange(2, colunaInicial, 1, 4)
    .setValues([cabecalho])
    .setFontWeight("bold")
    .setBackground("#000000") // Fundo Preto
    .setFontColor("#FFFFFF"); // Texto Branco

  if (!abaDados) return;

  // --- 4. LER DADOS ---
  const dados = abaDados.getDataRange().getDisplayValues();
  if (dados.length < 2) return;

  const linhas = dados.slice(1);
  const listaProcessada = [];

  // --- 5. PROCESSAR DADOS ---
  linhas.forEach(linha => {
    let mesCru = String(linha[0]); 
    const ano = linha[1];
    const vencimento = linha[2]; 
    const valorStr = linha[3]; 

    if (!mesCru || !ano) return;

    let mesLimpo = mesCru.trim();
    let indiceNum = 0;
    const match = mesCru.match(/\((\d+)\)/);
    if (match) {
      indiceNum = parseInt(match[1]); 
      mesLimpo = mesCru.replace(/\s*\(\d+\)/, "").trim(); 
    }
    mesLimpo = mesLimpo.toLowerCase();

    let valor = parseFloat(String(valorStr).replace("R$", "").replace(/\./g, "").replace(",", ".").trim()) || 0;

    listaProcessada.push({
      mes: mesLimpo,
      ano: ano,
      indice: indiceNum,
      vencimento: vencimento,
      valor: valor
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

  // --- 7. PREPARAR SAÍDA E CORES DE FUNDO ---
  const saida = [];
  const matrizFundos = []; 

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0); 

  listaProcessada.forEach(item => {
    const nomeMes = item.mes.charAt(0).toUpperCase() + item.mes.slice(1);
    const mesAno = `${nomeMes} ${item.ano}`;
    
    let tipo = item.indice === 0 ? "Padrão" : `Extra ${item.indice}`;
    
    let corFundo = "white"; // Padrão (Futuro)

    if (item.vencimento && item.vencimento.includes("/")) {
      const partes = item.vencimento.split("/"); 
      const dataConta = new Date(partes[2], partes[1] - 1, partes[0]);
      
      // Se já passou da data
      if (dataConta < hoje) {
        corFundo = "#EEEEEE"; // Cinza Claro
      }
    }

    saida.push([mesAno, tipo, item.vencimento, item.valor]);
    matrizFundos.push([corFundo, corFundo, corFundo, corFundo]);
  });

  // --- 8. ESCREVER DADOS ---
  // Começa na linha 3 (Título=1, Cabeçalho=2)
  const rangeTabela = abaRelatorio.getRange(3, colunaInicial, saida.length, 4);
  
  rangeTabela.setValues(saida);
  rangeTabela.setHorizontalAlignment("left");
  rangeTabela.setBackgrounds(matrizFundos); // Aplica o fundo cinza ou branco

  // Formatação R$
  abaRelatorio.getRange(3, colunaInicial + 3, saida.length, 1).setNumberFormat("R$ #,##0.00");
  
  abaRelatorio.autoResizeColumns(colunaInicial, 4);

  // --- 9. INSERIR RODAPÉ (DISCLAIMER) ---
  // Calcula a linha logo após o último dado
  const linhaRodape = 3 + saida.length; 

  const celulaRodape = abaRelatorio.getRange(linhaRodape, colunaInicial, 1, 4);
  celulaRodape
    .merge() // Mescla as 4 colunas
    .setValue("Linhas em cinza indicam que a data de vencimento já passou, mas não confirmam o pagamento.")
    .setFontSize(8)          // Letra menor
    .setFontStyle("italic")  // Itálico
    .setFontColor("#333") // Texto cinza escuro
    .setBackground("white")  // Fundo branco para destacar do resto
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setWrap(true);          // Quebra de texto se ficar muito longo
}