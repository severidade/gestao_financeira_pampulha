function gestao_pampulha_dados_brutos() {
  // --- CONFIGURAÇÕES ---
  const idPlanilhaOrigem = "1-AFYVWxgZZRK2mlHegka7RmepSbIOYWaw6ekc_GO4H8";
  const nomeAbaOrigem = "Respostas ao formulário 1";
  const nomeAbaDestino = "Dados_Brutos_2026";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDestino = ss.getSheetByName(nomeAbaDestino);

  if (!abaDestino) {
    SpreadsheetApp.getUi().alert(
      "Erro: Não encontrei a aba '" +
        nomeAbaDestino +
        "'. Verifique se criou esta aba na planilha.",
    );
    return;
  }

  try {
    const planilhaOrigem = SpreadsheetApp.openById(idPlanilhaOrigem);
    const abaOrigem = planilhaOrigem.getSheetByName(nomeAbaOrigem);

    if (!abaOrigem) {
      SpreadsheetApp.getUi().alert(
        "Erro: Não encontrei a aba de origem '" + nomeAbaOrigem + "'.",
      );
      return;
    }

    const ultimaLinha = abaOrigem.getLastRow();
    if (ultimaLinha < 2) return;

    let dados = abaOrigem
      .getRange(2, 1, ultimaLinha - 1, 17)
      .getDisplayValues();

    // ORDENAR POR DATA
    dados.sort(function (a, b) {
      const dataA = a[2];
      const dataB = b[2];
      function converterData(str) {
        if (!str) return 0;
        const partes = str.split("/");
        return new Date(partes[2], partes[1] - 1, partes[0]).getTime();
      }
      return converterData(dataA) - converterData(dataB);
    });

    let saidaRichText = [];

    // --- ESTILOS ---
    const estiloErro = SpreadsheetApp.newTextStyle()
      .setUnderline(false)
      .setForegroundColor("red")
      .build();
    const estiloNormal = SpreadsheetApp.newTextStyle()
      .setUnderline(false)
      .setForegroundColor("black")
      .build();

    // --- FUNÇÕES AJUDANTES ---
    function tratarValor(valorStr) {
      if (!valorStr) return 0;
      let numero = parseFloat(
        valorStr
          .toString()
          .replace(",", ".")
          .replace(/[^\d.-]/g, ""),
      );
      return isNaN(numero) ? 0 : numero;
    }

    function criarCelulasPagamento(pagou, valorStr, linkStr, dataStr) {
      let rtValor, rtData;
      let respostaPagou = pagou ? pagou.toString().trim().toLowerCase() : "";

      if (respostaPagou === "sim") {
        let valNum = tratarValor(valorStr);
        let textoVal = valNum.toLocaleString("pt-BR", {
          style: "currency",
          currency: "BRL",
        });

        let builderValor = SpreadsheetApp.newRichTextValue().setText(textoVal);

        if (linkStr && linkStr.toString().includes("http")) {
          builderValor.setLinkUrl(linkStr);
        } else {
          builderValor.setLinkUrl(null);
        }
        if (!linkStr) builderValor.setTextStyle(estiloNormal);
        rtValor = builderValor.build();

        if (dataStr && dataStr.toString().trim() !== "") {
          rtData = SpreadsheetApp.newRichTextValue()
            .setText(dataStr)
            .setTextStyle(estiloNormal)
            .setLinkUrl(null)
            .build();
        } else {
          rtData = SpreadsheetApp.newRichTextValue()
            .setText("🤬")
            .setTextStyle(estiloErro)
            .setLinkUrl(null)
            .build();
        }
      } else if (respostaPagou === "não") {
        let textoValZero = (0).toLocaleString("pt-BR", {
          style: "currency",
          currency: "BRL",
        });
        rtValor = SpreadsheetApp.newRichTextValue()
          .setText(textoValZero)
          .setTextStyle(estiloNormal)
          .setLinkUrl(null)
          .build();
        rtData = SpreadsheetApp.newRichTextValue()
          .setText("-")
          .setTextStyle(estiloNormal)
          .setLinkUrl(null)
          .build();
      } else {
        rtValor = SpreadsheetApp.newRichTextValue()
          .setText("")
          .setLinkUrl(null)
          .build();
        rtData = SpreadsheetApp.newRichTextValue()
          .setText("")
          .setLinkUrl(null)
          .build();
      }
      return [rtValor, rtData];
    }

    // --- LOOP PRINCIPAL ---
    dados.forEach((linha) => {
      let servico = linha[1];

      if (servico && servico.toString().trim() !== "") {
        let dtVencimento = linha[2];
        let valorBruto = linha[3];
        let linkDoc = linha[4];

        let valorNumerico = tratarValor(valorBruto);
        let textoValorFormatado = valorNumerico.toLocaleString("pt-BR", {
          style: "currency",
          currency: "BRL",
        });
        let valorPorPessoa = valorNumerico / 3;
        let textoDivisao = valorPorPessoa.toLocaleString("pt-BR", {
          style: "currency",
          currency: "BRL",
        });

        // QUEM PAGOU
        let pagadores = [];
        if (linha[5] && linha[5].trim().toLowerCase() === "sim")
          pagadores.push("Marco");
        if (linha[9] && linha[9].trim().toLowerCase() === "sim")
          pagadores.push("Janaína");
        if (linha[13] && linha[13].trim().toLowerCase() === "sim")
          pagadores.push("Adriana");
        let textoPagadores = pagadores.length > 0 ? pagadores.join(",") : "";

        // --- CONSTRUÇÃO DO ID (NOVO) ---
        // 1. Remove as barras da data (12/01/2026 -> 12012026)
        let dataSemBarra = dtVencimento.toString().split("/").join("");
        // 2. Cria o ID (12012026-COPASA) - UPPERCASE para evitar erros de busca
        let idUnico = dataSemBarra + "-" + servico.trim().toUpperCase();
        let rtId = SpreadsheetApp.newRichTextValue()
          .setText(idUnico)
          .setTextStyle(estiloNormal)
          .build();

        // COLUNAS BÁSICAS
        let rtVencimento = SpreadsheetApp.newRichTextValue()
          .setText(dtVencimento)
          .setTextStyle(estiloNormal)
          .setLinkUrl(null)
          .build();
        let rtServico = SpreadsheetApp.newRichTextValue()
          .setText(servico)
          .setTextStyle(estiloNormal)
          .setLinkUrl(null)
          .build();
        let rtValor = SpreadsheetApp.newRichTextValue()
          .setText(textoValorFormatado)
          .setTextStyle(estiloNormal)
          .setLinkUrl(null)
          .build();
        let rtQuem = SpreadsheetApp.newRichTextValue()
          .setText(textoPagadores)
          .setTextStyle(estiloNormal)
          .setLinkUrl(null)
          .build();

        let rtDoc;
        if (linkDoc && linkDoc.toString().includes("http")) {
          rtDoc = SpreadsheetApp.newRichTextValue()
            .setText("📄")
            .setLinkUrl(linkDoc)
            .build();
        } else {
          rtDoc = SpreadsheetApp.newRichTextValue()
            .setText("🤬")
            .setTextStyle(estiloErro)
            .setLinkUrl(null)
            .build();
        }

        let rtDivisao = SpreadsheetApp.newRichTextValue()
          .setText(textoDivisao)
          .setTextStyle(estiloNormal)
          .setLinkUrl(null)
          .build();
        let celulasMarco = criarCelulasPagamento(
          linha[5],
          linha[6],
          linha[8],
          linha[7],
        );
        let celulasJanaina = criarCelulasPagamento(
          linha[9],
          linha[10],
          linha[12],
          linha[11],
        );
        let celulasAdriana = criarCelulasPagamento(
          linha[13],
          linha[14],
          linha[16],
          linha[15],
        );

        // --- ORDEM DE SAÍDA (AQUI MUDOU) ---
        saidaRichText.push([
          rtId, // Coluna A (O ID VEM PRIMEIRO)
          rtVencimento, // Coluna B
          rtServico, // Coluna C
          rtValor, // Coluna D
          rtQuem, // Coluna E (Atenção: Validação deve vir pra cá)
          rtDoc, // Coluna F
          rtDivisao, // Coluna G
          celulasMarco[0],
          celulasMarco[1],
          celulasJanaina[0],
          celulasJanaina[1],
          celulasAdriana[0],
          celulasAdriana[1],
        ]);
      }
    });

    const totalLinhas = abaDestino.getMaxRows();
    // Limpa até a coluna 13 (A até M)
    if (totalLinhas > 1) {
      abaDestino.getRange(2, 1, totalLinhas - 1, 13).clearContent();
    }

    if (saidaRichText.length > 0) {
      abaDestino
        .getRange(2, 1, saidaRichText.length, 13)
        .setRichTextValues(saidaRichText);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro: " + e.message);
  }
}
