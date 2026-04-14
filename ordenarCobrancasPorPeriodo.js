function ordenarCobrancasPorPeriodo(mesTexto, anoTexto, indiceJaCalculado = null) {
  const mapaMeses = { 
    "janeiro":1, "fevereiro":2, "março":3, "marco":3, 
    "abril":4, "maio":5, "junho":6, 
    "julho":7, "agosto":8, "setembro":9, 
    "outubro":10, "novembro":11, "dezembro":12 
  };

  let mesLimpo = String(mesTexto).toLowerCase().trim();
  let indice = 0;

  // Se o índice não foi passado, tenta extrair do texto (ex: "janeiro (1)")
  if (indiceJaCalculado !== null) {
    indice = parseInt(indiceJaCalculado);
  } else {
    const match = mesLimpo.match(/\((\d+)\)/);
    if (match) {
      indice = parseInt(match[1]);
      mesLimpo = mesLimpo.split("(")[0].trim(); // Remove o (1) para achar no mapa
    }
  }

  const numMes = mapaMeses[mesLimpo] || 0;
  const numAno = parseInt(anoTexto) || 0;

  // Exemplo: 2026 + 01 + 01 = 20260101
  return (numAno * 10000) + (numMes * 100) + indice;
}