# Sistema Integrado de Gestão Financeira

Sistema integrado de gestão financeira destinado ao controle e rateio de despesas compartilhadas. A solução automatiza o ciclo financeiro completo desde o lançamento de passivos e auditoria de cotas até a notificação de cobrança e conciliação do fluxo de caixa.

## 🗂 Estrutura de Dados

Abaixo estão as definições das principais tabelas do sistema (clique para expandir):

<details>
  <summary><strong>💸 Passivos_Dados_Brutos</strong></summary>
  <br>
  É o "Banco de Dados de Despesas". Nesta tabela são lançados os serviços/passivos, seus valores e recibos.
</details>

<details>
  <summary><strong>🤝 Acertos_Mensais_Dados_Brutos</strong></summary>
  <br>
  É a "Tabela de Faturamento". Ela transforma a despesa (passivo) em cobrança (ativo). Contém o valor da cota individual por mês com os dados de pagamento do QR code.
</details>

<details>
  <summary><strong>💰 Fluxo_Caixa_Dados_Brutos</strong></summary>
  <br>
  É a "Tesouraria/Baixa". É a prova real de que o dinheiro saiu do bolso do morador e entrou na conta do gestor. Controla os valores recebidos, quem pagou, quando e quanto.
</details>

---

## 🔄 Fluxo de Entrada das Informações

O ciclo financeiro do sistema segue as etapas abaixo:

1.  **Entrada:** Chega a conta de Luz ➔ Você lança em *Passivos*.
2.  **Cálculo:** O mês fecha ➔ Você define o valor do rateio em *Acertos*.
3.  **Auditoria:** Você roda o Relatório de Conferência ➔ O script checa se *Passivos* bate com *Acertos*.
4.  **Cobrança:** Tudo certo? ➔ Você clica em *Enviar E-mail*.
5.  **Baixa:** O Pix cai na conta ➔ Você registra em *Fluxo de Caixa*.