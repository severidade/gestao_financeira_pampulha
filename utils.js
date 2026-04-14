function tratarValor(valorStr) {
  if (!valorStr) return 0;

  let str = valorStr.toString().trim();
  // Remove moeda e espaços
  str = str.replace(/[^\d,.-]/g, "");
  // Caso 1: formato BR → 1.234,56
  // Caso 2: formato BR simples → 26,09
  // Caso 3: formato EN → 26.09 👉 NÃO FAZ NADA
  if (str.includes(",") && str.includes(".")) {
    str = str.replace(/\./g, "").replace(",", ".");
  } else if (str.includes(",")) {
    str = str.replace(",", ".");
  }
  const numero = parseFloat(str);
  return isNaN(numero) ? 0 : numero;
}

function obterNumeroMes(nomeMes) {
  if (!nomeMes) return 0;
  const mesNormalizado = nomeMes
    .toString()
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, ""); // Remove cedilha (ex: março -> marco)

  const mapa = {
    janeiro: 1,
    fevereiro: 2,
    marco: 3,
    abril: 4,
    maio: 5,
    junho: 6,
    julho: 7,
    agosto: 8,
    setembro: 9,
    outubro: 10,
    novembro: 11,
    dezembro: 12,
  };

  return mapa[mesNormalizado] || 0;
}
