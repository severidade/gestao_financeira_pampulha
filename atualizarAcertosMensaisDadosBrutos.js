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

  // MEMÓRIA
  const dadosAtuais = abaDestino.getDataRange().getValues();
  const memoriaStatus = memoriaStatusEnvioCobrancaMensal(dadosAtuais);

  // ESTILOS
  const estilos = obterEstilosPlanilhas();

  // LIMPAR
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

  // CABEÇALHO
  const cabecalho = titulos.map((txt) =>
    render.criarTexto(txt, estilos.cabecalho),
  );
  saidaRichText.push(cabecalho);

  if (dadosOrigem.length >= 2) {
    let linhasDados = dadosOrigem.slice(1);

    // ORDENAÇÃO
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

    // PROCESSAMENTO
    linhasDados.forEach((linha) => {
      const mesRef = linha[1];
      const anoRef = linha[2];
      const valorBruto = linha[3];
      const dataVencimento = linha[4];
      const inputSuplementar = linha[5];
      const linkQrOriginal = linha[6];
      const chavePix = linha[7];

      if (mesRef || valorBruto) {
        const valorNumerico = tratarValor(valorBruto);

        let textoMesComposto = mesRef;
        let indiceSup = 0;

        if (inputSuplementar && inputSuplementar.toString().trim() !== "") {
          indiceSup = parseInt(inputSuplementar.toString().trim());
        }

        if (indiceSup > 0) {
          textoMesComposto = `${mesRef} (${indiceSup})`;
        }

        const chaveAtual = `${textoMesComposto}|${anoRef}`;

        let statusParaGravar = "-";
        if (chaveAtual in memoriaStatus) {
          statusParaGravar = memoriaStatus[chaveAtual];
        }

        const textoValor = valorNumerico.toLocaleString("pt-BR", {
          style: "currency",
          currency: "BRL",
        });

        const rtQr = render.criarLink(linkQrOriginal, estilos);

        saidaRichText.push([
          render.criarTexto(textoMesComposto, estilos.normal),
          render.criarTexto(anoRef, estilos.normal),
          render.criarTexto(dataVencimento || "-", estilos.normal),
          render.criarTexto(textoValor, estilos.normal),
          rtQr,
          render.criarTexto(chavePix || "-", estilos.normal),
          render.criarTexto(statusParaGravar, estilos.normal),
        ]);

        valoresNumericos.push([valorNumerico]);
      }
    });
  }

  // ESCRITA FINAL
  if (saidaRichText.length > 0) {
    escreverTabelaAcertoMensal(
      abaDestino,
      saidaRichText,
      valoresNumericos,
      estilos,
    );
  }
}
