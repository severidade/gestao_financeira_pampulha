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

  // ESTILOS
  const estilos = obterEstilosPlanilhas();

  // LIMPAR
  abaDestino.clear();
  SpreadsheetApp.flush();

  const dadosOrigem = abaOrigem.getDataRange().getDisplayValues();

  const saidaRichText = [];

  const valoresNumericos = [];

  const titulos = ["Mês Ref.", "Ano", "Serviço", "Valor", "Recibo", "Pago por"];

  const cabecalho = titulos.map((txt) =>
    render.criarTexto(txt, estilos.cabecalho),
  );
  saidaRichText.push(cabecalho);

  if (dadosOrigem.length >= 2) {
    let linhasDados = dadosOrigem.slice(1);

    // ORDENAÇÃO
    linhasDados.sort(function (a, b) {
      const anoA = parseInt(a[3]) || 0;
      const anoB = parseInt(b[3]) || 0;

      if (anoA !== anoB) return anoA - anoB;

      const mesA = obterNumeroMes(a[2]);
      const mesB = obterNumeroMes(b[2]);

      return mesA - mesB;
    });

    // PROCESSAMENTO
    linhasDados.forEach((linha) => {
      const servico = linha[1];
      let mesRef = linha[2];
      const anoRef = linha[3];
      const valorBruto = linha[5];
      const linkDoc = linha[7];
      const numSuplementar = linha[4];
      const quemPagou = linha[8] || "Rateio automático";

      if (servico || valorBruto) {
        if (numSuplementar) {
          mesRef = `${mesRef} (${numSuplementar})`;
        }

        const valorNumerico = tratarValor(valorBruto);

        const textoValor = valorNumerico.toLocaleString("pt-BR", {
          style: "currency",
          currency: "BRL",
        });

        // link de recibo
        const rtRecibo = render.criarLinkRecibo(linkDoc, estilos);

        saidaRichText.push([
          render.criarTexto(mesRef, estilos.normal),
          render.criarTexto(anoRef, estilos.normal),
          render.criarTexto(servico || "-", estilos.normal),
          render.criarTexto(textoValor, estilos.normal),
          rtRecibo,
          render.criarTexto(quemPagou, estilos.normal),
        ]);

        valoresNumericos.push([valorNumerico]);
      }
    });
  }

  // ESCRITA FINAL
  if (saidaRichText.length > 0) {
    escreverTabelaPassivos(
      abaDestino,
      saidaRichText,
      valoresNumericos,
      estilos,
    );
  }
}


// MOdelo do array bidimensional
// [
//   ["Mês", "Ano", "Valor"],
//   ["Jan", "2024", "100"],
//   ["Fev", "2024", "200"]
// ]
