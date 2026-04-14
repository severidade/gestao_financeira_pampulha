function processarEntradaFluxoCaixa(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomeAba = "💰 Fluxo_Caixa_Dados_Brutos";
  let aba = ss.getSheetByName(nomeAba);

  const ID_PASTA_COMPROVANTES = "1_GVM4JOeeRaa2h7wjFF6W3KIXVWJaBT1";

  // --- 1. GARANTIA DA ABA E CABEÇALHO ---
  if (!aba) {
    aba = ss.insertSheet(nomeAba);
  }

  aba
    .getRange(1, 1, 1, 8)
    .setValues([
      [
        "Mês Ref",
        "Ano",
        "Marco (Valor)",
        "Data Pag.",
        "Janaína (Valor)",
        "Data Pag.",
        "Adriana (Valor)",
        "Data Pag.",
      ],
    ]);
  aba.setFrozenRows(1);
  aba
    .getRange("A1:H1")
    .setFontWeight("bold")
    .setBackground("#000000")
    .setFontColor("#ffffff");

  // --- 2. IDENTIFICAR LINHA ---
  let dadosSheet = aba.getDataRange().getValues();
  let linhaEncontrada = -1;

  let mesPayload = String(dados.mesRef || "").trim();
  let anoPayload = String(dados.anoRef || "").trim();

  console.log(`🔎 Procurando por: Mês [${mesPayload}] e Ano [${anoPayload}]`);

  for (let i = 1; i < dadosSheet.length; i++) {
    let mesLinha = String(dadosSheet[i][0]).trim();
    let anoLinha = String(dadosSheet[i][1]).trim();

    if (mesLinha === mesPayload && anoLinha === anoPayload) {
      linhaEncontrada = i + 1;
      console.log(`✅ Linha encontrada: ${linhaEncontrada}`);
      break;
    }
  }

  // --- 3. SE NÃO ACHAR, CRIA LINHA ---
  if (linhaEncontrada === -1) {
    if (!anoPayload || anoPayload === "undefined") {
      anoPayload = new Date().getFullYear().toString();
    }

    aba.appendRow([mesPayload, anoPayload, 0, "😢", 0, "😢", 0, "😢"]);
    linhaEncontrada = aba.getLastRow();
    console.log(`🆕 Nova linha criada: ${linhaEncontrada}`);
  }

  // --- 4. MAPA DAS COLUNAS ---
  const mapaColunas = {
    Marco: { colValor: 3, colData: 4 },
    Janaina: { colValor: 5, colData: 6 },
    Adriana: { colValor: 7, colData: 8 },
  };

  const estiloLink = SpreadsheetApp.newTextStyle()
    .setUnderline(true)
    .setForegroundColor("#1155cc")
    .build();
  const estiloNormal = SpreadsheetApp.newTextStyle()
    .setUnderline(false)
    .setForegroundColor("#000000")
    .build();

  // --- 5. GRAVAR DADOS ---
  if (!dados.pagamentos || dados.pagamentos.length === 0) {
    console.warn("⚠️ Nenhum pagamento recebido.");
  }

  dados.pagamentos.forEach((pgto) => {
    let nomeChave = pgto.pessoa
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "");
    if (nomeChave.toLowerCase().includes("marco")) nomeChave = "Marco";
    else if (nomeChave.toLowerCase().includes("janaina")) nomeChave = "Janaina";
    else if (nomeChave.toLowerCase().includes("adriana")) nomeChave = "Adriana";

    const coords = mapaColunas[nomeChave];

    if (coords) {
      // Data
      let dataParts = pgto.data.split("-");
      let dataObj = new Date(dataParts[0], dataParts[1] - 1, dataParts[2]);
      let dataFormatada = Utilities.formatDate(
        dataObj,
        ss.getSpreadsheetTimeZone(),
        "dd/MM",
      );

      // Valor (Tratamento Blindado)
      let valorTexto = String(pgto.valor).replace("R$", "").trim();
      if (valorTexto.indexOf(",") > -1 && valorTexto.indexOf(".") > -1) {
        valorTexto = valorTexto.replace(/\./g, "");
      }
      valorTexto = valorTexto.replace(",", ".");
      let valorLimpo = parseFloat(valorTexto);
      if (isNaN(valorLimpo)) valorLimpo = 0;

      // Grava Valor
      aba.getRange(linhaEncontrada, coords.colValor).setValue(valorLimpo);

      // Grava Arquivo / Data
      let valorCelulaData = null;
      if (pgto.arquivo) {
        try {
          console.log(`📂 Salvando arquivo para ${nomeChave}...`);
          let blob = Utilities.newBlob(
            Utilities.base64Decode(pgto.arquivo.base64),
            pgto.arquivo.tipo,
            `${mesPayload}_${anoPayload}_${nomeChave}_comprovante`,
          );

          let pasta;
          try {
            pasta = DriveApp.getFolderById(ID_PASTA_COMPROVANTES);
          } catch (e) {
            console.error("Erro ID pasta, salvando na raiz.");
            pasta = DriveApp.getRootFolder();
          }

          let arquivoDrive = pasta.createFile(blob);

          valorCelulaData = SpreadsheetApp.newRichTextValue()
            .setText(dataFormatada)
            .setLinkUrl(arquivoDrive.getUrl())
            .setTextStyle(estiloLink)
            .build();
        } catch (e) {
          console.error("❌ Erro arquivo:", e);
          valorCelulaData = SpreadsheetApp.newRichTextValue()
            .setText(dataFormatada)
            .setTextStyle(estiloNormal)
            .build();
        }
      } else {
        valorCelulaData = SpreadsheetApp.newRichTextValue()
          .setText(dataFormatada)
          .setLinkUrl(null)
          .setTextStyle(estiloNormal)
          .build();
      }

      aba
        .getRange(linhaEncontrada, coords.colData)
        .setRichTextValue(valorCelulaData);
    }
  });

  // --- 6. FORMATAÇÃO ---
  aba.getRange(linhaEncontrada, 3).setNumberFormat("R$ #,##0.00");
  aba.getRange(linhaEncontrada, 5).setNumberFormat("R$ #,##0.00");
  aba.getRange(linhaEncontrada, 7).setNumberFormat("R$ #,##0.00");
  aba.getRange(linhaEncontrada, 1, 1, 8).setHorizontalAlignment("center");

  // 🛡️ GARANTE PERSISTÊNCIA (sem ordenação agora)
  SpreadsheetApp.flush();
}
