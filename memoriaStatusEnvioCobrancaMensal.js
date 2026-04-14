function memoriaStatusEnvioCobrancaMensal(dadosAtuais) {
  const memoriaStatus = {};

  if (dadosAtuais.length > 1) {
    for (let i = 1; i < dadosAtuais.length; i++) {
      const mes = String(dadosAtuais[i][0]).trim(); // Coluna A (Ex: Janeiro (1))
      const ano = String(dadosAtuais[i][1]).trim(); // Coluna B (Ex: 2026)
      const status = dadosAtuais[i][6]; // Coluna G (Onde está escrito "✅ Enviado em 18/03/2026 21:46")

      if (mes && ano && status !== "") {
        memoriaStatus[`${mes}|${ano}`] = status;
      }
    }
  }

  return memoriaStatus;
}

// ============================================================
// 🛡️ PASSO 1: ATIVAR A MEMÓRIA (ANTES DE APAGAR)
//
// Contexto:
// Na aba "🤝 Acertos_Mensais_Dados_Brutos", a coluna "Status de envio"
// indica se a cobrança já foi encaminhada por e-mail.
//
// Esse status é atualizado localmente por um script, enquanto os demais
// dados da tabela são provenientes de uma planilha externa (formulário)
// e, portanto, são tratados como dados estáticos.
//
// Problema:
// Ao atualizar a tabela, todos os dados são apagados e recriados,
// o que faria com que o "Status de envio" fosse perdido.
//
// Solução:
// Antes de limpar a aba, os status existentes são armazenados em memória.
// Após a reconstrução da tabela, esses valores são restaurados,
// preservando a informação dinâmica.
// ============================================================
