function gastos_tabela() {
  // --- CONFIGURAÇÕES ---
  const idPlanilhaOrigem = "1e9cFyAkLPIQBl6Omn_ss7BQuYSpHSFtUcGHsoO5m8mM";
  const nomeAbaOrigem = "Respostas ao formulário 1";
  const nomeAbaDestino = "2026_tabela";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDestino = ss.getSheetByName(nomeAbaDestino);

  if (!abaDestino) {
    SpreadsheetApp.getUi().alert(
      "Erro: Não encontrei a aba '" + nomeAbaDestino + "'.",
    );
    return;
  }

  try {
    const planilhaOrigem = SpreadsheetApp.openById(idPlanilhaOrigem);
    const abaOrigem = planilhaOrigem.getSheetByName(nomeAbaOrigem);

    const ultimaLinha = abaOrigem.getLastRow();
    if (ultimaLinha < 2) return;

    // --- ALTERAÇÃO: Agora pegamos 17 colunas para alcançar as datas de pagamento no final
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

    // --- FUNÇÃO AJUDANTE PARA TRATAR VALORES ---
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

    // --- FUNÇÃO AJUDANTE PARA CRIAR CÉLULA DE PAGAMENTO INDIVIDUAL ---
    function criarCelulasPagamento(pagou, valorStr, linkStr, dataStr) {
      let rtValor, rtData;

      // Se pagou "Sim"
      if (pagou && pagou.toString().trim().toLowerCase() === "sim") {
        // Formata Valor
        let valNum = tratarValor(valorStr);
        let textoVal = valNum.toLocaleString("pt-BR", {
          style: "currency",
          currency: "BRL",
        });

        let builderValor = SpreadsheetApp.newRichTextValue().setText(textoVal);

        // Adiciona Link se existir
        if (linkStr && linkStr.toString().includes("http")) {
          builderValor.setLinkUrl(linkStr);
        } else {
          builderValor.setLinkUrl(null);
        }
        rtValor = builderValor.build();

        // Data do Pagamento
        rtData = SpreadsheetApp.newRichTextValue()
          .setText(dataStr ? dataStr : "-") // Se não tiver data, põe traço
          .setLinkUrl(null)
          .build();
      } else {
        // Se não pagou ou vazio
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
        // 1. DADOS ORIGINAIS
        let dtVencimento = linha[2];
        let valorBruto = linha[3];
        let linkDoc = linha[4];

        // Tratamento Valor Geral
        let valorNumerico = tratarValor(valorBruto);
        let textoValorFormatado = valorNumerico.toLocaleString("pt-BR", {
          style: "currency",
          currency: "BRL",
        });

        // 2. RATEIO (Coluna F)
        let valorPorPessoa = valorNumerico / 3;
        let textoDivisao = valorPorPessoa.toLocaleString("pt-BR", {
          style: "currency",
          currency: "BRL",
        });

        // 3. QUEM PAGOU (Coluna D - Texto simples)
        let pagadores = [];
        if (linha[5] && linha[5].trim().toLowerCase() === "sim")
          pagadores.push("Marco");
        if (linha[8] && linha[8].trim().toLowerCase() === "sim")
          pagadores.push("Janaína");
        if (linha[11] && linha[11].trim().toLowerCase() === "sim")
          pagadores.push("Adriana");
        let textoPagadores = pagadores.length > 0 ? pagadores.join(",") : "";

        // --- CONSTRUÇÃO DAS COLUNAS BÁSICAS ---
        let rtVencimento = SpreadsheetApp.newRichTextValue()
          .setText(dtVencimento)
          .setLinkUrl(null)
          .build();
        let rtServico = SpreadsheetApp.newRichTextValue()
          .setText(servico)
          .setLinkUrl(null)
          .build();
        let rtValor = SpreadsheetApp.newRichTextValue()
          .setText(textoValorFormatado)
          .setLinkUrl(null)
          .build();
        let rtQuem = SpreadsheetApp.newRichTextValue()
          .setText(textoPagadores)
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
            .setText("🤬 putz! nao mandou o recibo")
            .setTextStyle(estiloErro)
            .setLinkUrl(null)
            .build();
        }

        let rtDivisao = SpreadsheetApp.newRichTextValue()
          .setText(textoDivisao)
          .setLinkUrl(null)
          .build();

        // --- NOVAS COLUNAS: MARCO, JANAÍNA, ADRIANA ---
        // Índices baseados no cabeçalho fornecido:
        // Marco: Pagou?(5), Valor(6), Link(7), Data(14)
        let celulasMarco = criarCelulasPagamento(
          linha[5],
          linha[6],
          linha[7],
          linha[14],
        );

        // Janaína: Pagou?(8), Valor(9), Link(10), Data(15)
        let celulasJanaina = criarCelulasPagamento(
          linha[8],
          linha[9],
          linha[10],
          linha[15],
        );

        // Adriana: Pagou?(11), Valor(12), Link(13), Data(16)
        let celulasAdriana = criarCelulasPagamento(
          linha[11],
          linha[12],
          linha[13],
          linha[16],
        );

        // MONTAGEM DA LINHA (12 Colunas)
        saidaRichText.push([
          rtVencimento, // A
          rtServico, // B
          rtValor, // C
          rtQuem, // D
          rtDoc, // E
          rtDivisao, // F
          celulasMarco[0], // G (Marco Valor com Link)
          celulasMarco[1], // H (Marco Data)
          celulasJanaina[0], // I (Janaína Valor com Link)
          celulasJanaina[1], // J (Janaína Data)
          celulasAdriana[0], // K (Adriana Valor com Link)
          celulasAdriana[1], // L (Adriana Data)
        ]);
      }
    });

    // --- LIMPEZA (AGORA COM 12 COLUNAS) ---
    const totalLinhas = abaDestino.getMaxRows();
    if (totalLinhas > 1) {
      abaDestino.getRange(2, 1, totalLinhas - 1, 12).clearContent();
    }

    // --- ESCRITA (AGORA COM 12 COLUNAS) ---
    if (saidaRichText.length > 0) {
      abaDestino
        .getRange(2, 1, saidaRichText.length, 12)
        .setRichTextValues(saidaRichText);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro: " + e.message);
  }
}
