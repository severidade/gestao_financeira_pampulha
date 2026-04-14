function isContaVencida(dataVencimento) {
  // Segurança: Se não for uma data válida, não consideramos vencida (para não sumir com o dado)
  if (!dataVencimento || !(dataVencimento instanceof Date)) {
    return false;
  }

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0); // Zera hora, minuto, segundo de hoje

  const dataComparacao = new Date(dataVencimento);
  dataComparacao.setHours(0, 0, 0, 0); // Zera hora da data alvo

  // Se a data do vencimento for MENOR que hoje, já era. Venceu.
  return dataComparacao < hoje;
}