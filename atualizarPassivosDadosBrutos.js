function atualizarPassivosDadosBrutos() {
  const idPlanilhaOrigem = "1cpMvXFqYFZAPQ-7ek1bLjRamwjJ8JhCO3pXaGcyHfA8";
  const nomeAbaOrigem = "Respostas ao formulário 1";
  const nomeAbaDestino = "💸 Passivos_Dados_Brutos";

  const ssOrigem = SpreadsheetApp.openById(idPlanilhaOrigem);
  const abaOrigem = ssOrigem.getSheetByName(nomeAbaOrigem);

  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  let abaDestino = ssDestino.getSheetByName(nomeAbaDestino);

  if (!abaDestino) {
    abaDestino = ssDestino.insertSheet(nomeAbaDestino);
  }

  // --- Sempre que Atualizar ele limpa a aga destino ---
  abaDestino.clear();

  // --- Pega tudo na aba origem e transforma em string --- //
  const dadosOrigem = abaOrigem.getDataRange().getDisplayValues();
  const saida = [];

  // Cria o cabeçalho do meu array bidimensional //
  saida.push(["Mês Ref.", "Ano", "Serviço", "Valor", "Recibo", "Pago por"]);

  if (dadosOrigem.length >= 2) {
    let linhasDados = dadosOrigem.slice(1);

    // --- ORDENAÇÃO ---
    linhasDados.sort(function (a, b) {
      const anoA = parseInt(a[3]) || 0;
      const anoB = parseInt(b[3]) || 0;

      const mesA = obterNumeroMes(a[2]);
      const mesB = obterNumeroMes(b[2]);

      if (anoA !== anoB) return anoA - anoB;
      return mesA - mesB;
    });

    // --- PROCESSAMENTO ---
    linhasDados.forEach((linha) => {
      const servico = linha[1];
      let mesRef = linha[2];
      const anoRef = linha[3];
      const valorBruto = linha[5];
      const linkDoc = linha[7];
      const numSuplementar = linha[4];
      const quemPagou = linha[8] || "Rateio automático";

      if (servico || valorBruto) {
        if (numSuplementar && numSuplementar.toString().trim() !== "") {
          mesRef = mesRef + " (" + numSuplementar.toString().trim() + ")";
        }

        const valorNumerico = tratarValor(valorBruto);

        saida.push([
          mesRef,
          anoRef,
          servico,
          valorNumerico,
          linkDoc,
          quemPagou,
        ]);
      }
    });
  }

  // --- ESCREVER DADOS ---
  if (saida.length > 1) {
    const numLinhas = saida.length;

    const range = abaDestino.getRange(1, 1, numLinhas, 6);
    range.setValues(saida);

    // --- FORMATAR VALORES ---
    abaDestino.getRange(2, 4, numLinhas - 1, 1).setNumberFormat("R$ #,##0.00");

    // --- TRANSFORMAR LINKS EM 📄 CLICÁVEL ---
    const colunaRecibos = 5;

    for (let i = 1; i < saida.length; i++) {
      const link = saida[i][4];

      if (link && link.toString().includes("http")) {
        const richText = SpreadsheetApp.newRichTextValue()
          .setText("📄")
          .setLinkUrl(link)
          .build();

        abaDestino.getRange(i + 1, colunaRecibos).setRichTextValue(richText);
      } else {
        abaDestino.getRange(i + 1, colunaRecibos).setValue("🤬");
      }
    }

    abaDestino.autoResizeColumns(1, 6);
  }
}

// MOdelo do array bidimensional
// [
//   ["Mês", "Ano", "Valor"],
//   ["Jan", "2024", "100"],
//   ["Fev", "2024", "200"]
// ]
