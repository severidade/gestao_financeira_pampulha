function gestao_pampulha_passivos_dados_brutos() {

  const idPlanilhaOrigem = "1AvkuhDY96VOJTU63MiuizmGFLye0_0F2lpqvGsgx85k"; 
  const nomeAbaOrigem = "Respostas ao formulário 1"; 
  const nomeAbaDestino = "💸 Passivos_Dados_Brutos"; 

  const ssOrigem = SpreadsheetApp.openById(idPlanilhaOrigem);
  const abaOrigem = ssOrigem.getSheetByName(nomeAbaOrigem);
  
  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  let abaDestino = ssDestino.getSheetByName(nomeAbaDestino);

  if (!abaDestino) {
    abaDestino = ssDestino.insertSheet(nomeAbaDestino);
  }

  const estiloErro = SpreadsheetApp.newTextStyle().setUnderline(false).setForegroundColor("red").build();
  const estiloNormal = SpreadsheetApp.newTextStyle().setUnderline(false).setForegroundColor("black").build();

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

  abaDestino.clear();
  const dadosOrigem = abaOrigem.getDataRange().getDisplayValues(); 
  
  const saidaRichText = [];
  const valoresNumericos = [["Valor"]];

  // 🔥 Cabeçalho agora com coluna "Pago por"
  const cabecalho = ["Mês Ref.", "Ano", "Serviço", "Valor", "Recibo", "Pago por"]
    .map(txt => SpreadsheetApp.newRichTextValue().setText(txt).setTextStyle(estiloNormal).build());

  saidaRichText.push(cabecalho);

  if (dadosOrigem.length >= 2) {

    let linhasDados = dadosOrigem.slice(1);

    linhasDados.sort(function(a, b) {
      const anoA = parseInt(a[3]) || 0;
      const anoB = parseInt(b[3]) || 0;
      const mesA = obterNumeroMes(a[2]);
      const mesB = obterNumeroMes(b[2]);
      if (anoA !== anoB) return anoA - anoB; 
      return mesA - mesB;   
    });

    linhasDados.forEach(linha => {

      const servico = linha[1]; 
      let mesRef = linha[2]; 
      const anoRef = linha[3];
      const valorBruto = linha[4]; 
      const linkDoc = linha[6]; 
      const numSuplementar = linha[7];
      const quemPagou = linha[8] || "Não paga (rateio automático)";

      if (servico || valorBruto) {
        
        if (numSuplementar && numSuplementar.toString().trim() !== "") {
           mesRef = mesRef + " (" + numSuplementar.toString().trim() + ")";
        }

        let valorNumerico = tratarValor(valorBruto); 

        let rtDoc;
        if (linkDoc && linkDoc.toString().includes("http")) {
          rtDoc = SpreadsheetApp.newRichTextValue().setText("📄").setLinkUrl(linkDoc).build();
        } else {
          rtDoc = SpreadsheetApp.newRichTextValue().setText("🤬").setTextStyle(estiloErro).build();
        }

        let rtMes = SpreadsheetApp.newRichTextValue().setText(mesRef).setTextStyle(estiloNormal).build();
        let rtAno = SpreadsheetApp.newRichTextValue().setText(anoRef).setTextStyle(estiloNormal).build();
        let rtServico = SpreadsheetApp.newRichTextValue().setText(servico).setTextStyle(estiloNormal).build();
        let rtValor = SpreadsheetApp.newRichTextValue().setText(String(valorNumerico)).setTextStyle(estiloNormal).build();
        let rtPagoPor = SpreadsheetApp.newRichTextValue().setText(quemPagou).setTextStyle(estiloNormal).build();

        saidaRichText.push([rtMes, rtAno, rtServico, rtValor, rtDoc, rtPagoPor]);
        valoresNumericos.push([valorNumerico]);
      }
    });
  }

  if (saidaRichText.length > 0) {
    const numLinhas = saidaRichText.length;

    abaDestino.getRange(1, 4, numLinhas, 1).setNumberFormat("R$ #,##0.00");
    abaDestino.getRange(1, 1, numLinhas, 6).setRichTextValues(saidaRichText);

    if(valoresNumericos.length > 1){
      abaDestino.getRange(1, 4, valoresNumericos.length, 1).setValues(valoresNumericos);
    }

    abaDestino.autoResizeColumns(1, 6);
  }
}