function atualizarAcertosMensaisDadosBrutos() {
  const idPlanilhaOrigem = "1Jyt5_9ooJKgvsMALJDoWQB9cZkFrTHi0jkKMm0jVvcc";
  const nomeAbaOrigem = "Respostas ao formulário 1";
  const nomeAbaDestino = "🤝 Acertos_Mensais_Dados_Brutos";

  const ssOrigem = SpreadsheetApp.openById(idPlanilhaOrigem);
  const abaOrigem = ssOrigem.getSheetByName(nomeAbaOrigem);

  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  let abaDestino = ssDestino.getSheetByName(nomeAbaDestino);

  if (!abaDestino) {
    abaDestino = ssDestino.insertSheet(nomeAbaDestino);
  }

  // ============================================================
  // 🛡️ PASSO 1: ATIVAR A MEMÓRIA (ANTES DE APAGAR)
  // Nesta aba de acertos Mensais é gravado o status de envio que serve para identificar se uma cobrança foi enviada por e-mail
  // Por esse motivo antes de atualizar a tabela salvo na constante memoriaStatus o s valores
  // ============================================================
  const memoriaStatus = {};

  // Pega todos os dados que estão na planilha AGORA
  const dadosAtuais = abaDestino.getDataRange().getValues();

  // Se tiver dados (mais que 1 linha), vamos memorizar
  if (dadosAtuais.length > 1) {
    for (let i = 1; i < dadosAtuais.length; i++) {
      let mesChave = String(dadosAtuais[i][0]).trim(); // Coluna A (Ex: Janeiro (1))
      let anoChave = String(dadosAtuais[i][1]).trim(); // Coluna B (Ex: 2026)
      let status = dadosAtuais[i][6]; // Coluna G (Onde você escreveu ✅ Pago)

      // Se tiver algo escrito na Coluna G, guarda no "bolso" do script
      if (mesChave && anoChave && status !== "") {
        let chaveUnica = `${mesChave}|${anoChave}`;
        memoriaStatus[chaveUnica] = status;
      }
    }
  }
  // ============================================================

  // --- ESTILOS ---
  const estiloCabecalho = SpreadsheetApp.newTextStyle()
    .setFontFamily("Jost")
    .setUnderline(false)
    .setForegroundColor("wite")
    .build();
  const estiloNormal = SpreadsheetApp.newTextStyle()
    .setFontFamily("Lato")
    .setUnderline(false)
    .setForegroundColor("black")
    .build();
  const estiloLink = SpreadsheetApp.newTextStyle()
    .setUnderline(true)
    .setForegroundColor("#1155cc")
    .build();

  // --- PASSO 2: APAGAR TUDO (Agora é seguro, pois já memorizamos) ---
  abaDestino.clear();
  SpreadsheetApp.flush();

  const dadosOrigem = abaOrigem.getDataRange().getDisplayValues();

  const saidaRichText = [];
  const valoresNumericos = [["Valor"]];

  const titulos = [
    "Mês Ref.",
    "Ano",
    "Vencimento",
    "Valor",
    "QR Code",
    "Chave Pix",
    "Status Envio",
  ];

  // aqui
  const cabecalho = titulos.map((txt) =>
    SpreadsheetApp.newRichTextValue()
      .setText(txt)
      .setTextStyle(estiloCabecalho)
      .build(),
  );
  saidaRichText.push(cabecalho);

  if (dadosOrigem.length >= 2) {
    let linhasDados = dadosOrigem.slice(1);

    linhasDados.sort(function (a, b) {
      const anoA = parseInt(a[2]) || 0;
      const anoB = parseInt(b[2]) || 0;
      if (anoA !== anoB) return anoA - anoB;

      const mesA = obterNumeroMes(a[1]);
      const mesB = obterNumeroMes(b[1]);
      if (mesA !== mesB) return mesA - mesB;

      const supA = parseInt(a[6]) || 0;
      const supB = parseInt(b[6]) || 0;
      return supA - supB;
    });

    linhasDados.forEach((linha) => {
      const mesRef = linha[1];
      const anoRef = linha[2];
      const valorBruto = linha[3];
      const dataVencimento = linha[4];
      const inputSuplementar = linha[5];
      const linkQrOriginal = linha[6];
      const chavePix = linha[7];

      if (mesRef || valorBruto) {
        let valorNumerico = tratarValor(valorBruto);

        let textoMesComposto = mesRef;
        let indiceSup = 0;
        if (inputSuplementar && inputSuplementar.toString().trim() !== "") {
          indiceSup = parseInt(inputSuplementar.toString().trim());
        }
        if (indiceSup > 0) {
          textoMesComposto = `${mesRef} (${indiceSup})`;
        }

        let rtQr;
        if (linkQrOriginal && linkQrOriginal.toString().includes("http")) {
          rtQr = SpreadsheetApp.newRichTextValue()
            .setText("📱 Abrir")
            .setLinkUrl(linkQrOriginal)
            .setTextStyle(estiloLink)
            .build();
        } else {
          rtQr = SpreadsheetApp.newRichTextValue()
            .setText("-")
            .setTextStyle(estiloNormal)
            .build();
        }

        // ========================================================
        // 🛡️ PASSO 3: RESTAURAR O STATUS
        // ========================================================
        let statusParaGravar = "-";
        let chaveAtual = `${textoMesComposto}|${anoRef}`;

        // Verifica se temos algo guardado para este Mês/Ano
        if (memoriaStatus[chaveAtual]) {
          statusParaGravar = memoriaStatus[chaveAtual];
        }
        // ========================================================

        let textoVencimento = dataVencimento || "-";

        let rtMes = SpreadsheetApp.newRichTextValue()
          .setText(textoMesComposto)
          .setTextStyle(estiloNormal)
          .build();
        let rtAno = SpreadsheetApp.newRichTextValue()
          .setText(anoRef)
          .setTextStyle(estiloNormal)
          .build();
        let rtVencimento = SpreadsheetApp.newRichTextValue()
          .setText(textoVencimento)
          .setTextStyle(estiloNormal)
          .build();
        let rtChave = SpreadsheetApp.newRichTextValue()
          .setText(chavePix || "-")
          .setTextStyle(estiloNormal)
          .build();
        let textoValor = valorNumerico.toLocaleString("pt-BR", {
          style: "currency",
          currency: "BRL",
        });
        let rtValor = SpreadsheetApp.newRichTextValue()
          .setText(textoValor)
          .setTextStyle(estiloNormal)
          .build();

        // Coluna G: Grava o que recuperamos da memória
        let rtStatus = SpreadsheetApp.newRichTextValue()
          .setText(statusParaGravar)
          .setTextStyle(estiloNormal)
          .build();

        saidaRichText.push([
          rtMes,
          rtAno,
          rtVencimento,
          rtValor,
          rtQr,
          rtChave,
          rtStatus,
        ]);
        valoresNumericos.push([valorNumerico]);
      }
    });
  }

  // --- ESCREVER NA PLANILHA ---
  if (saidaRichText.length > 0) {
    const numLinhas = saidaRichText.length;

    abaDestino.getRange(1, 1, numLinhas, 7).setRichTextValues(saidaRichText);
    abaDestino.getRange(1, 4, numLinhas, 1).setNumberFormat("R$ #,##0.00");

    if (valoresNumericos.length > 0) {
      abaDestino
        .getRange(1, 4, valoresNumericos.length, 1)
        .setValues(valoresNumericos);
    }

    abaDestino.getRange(1, 1, numLinhas, 7).setHorizontalAlignment("center");
    abaDestino.getRange(1, 6, numLinhas, 1).setHorizontalAlignment("left");

    abaDestino.autoResizeColumns(1, 7);
  }
}
