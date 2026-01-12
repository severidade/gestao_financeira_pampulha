function notificar_novo_acerto() {
  // --- 1. CONFIGURAÇÃO ---
  const listaDestinatarios = [
    "oliveira.severo@gmail.com" 
  ];

  // --- 2. ACESSAR DADOS DO ACERTO (ÚLTIMA LINHA) ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAcertos = ss.getSheetByName("🤝 Acertos_Mensais_Dados_Brutos");
  const abaPassivos = ss.getSheetByName("💸 Passivos_Dados_Brutos");
  const ui = SpreadsheetApp.getUi();

  const ultimaLinha = abaAcertos.getLastRow();
  if (ultimaLinha < 2) {
    ui.alert("A tabela de acertos está vazia.");
    return;
  }

  // Pega dados do Acerto
  const mesAcerto = abaAcertos.getRange(ultimaLinha, 1).getValue();
  const anoAcerto = abaAcertos.getRange(ultimaLinha, 2).getValue();
  const valorFormatado = abaAcertos.getRange(ultimaLinha, 3).getDisplayValue(); 
  const chavePix = abaAcertos.getRange(ultimaLinha, 5).getValue(); 

  // --- 3. BUSCAR DETALHES E ANEXOS (PASSIVOS) ---
  const rangePassivos = abaPassivos.getDataRange();
  const dadosPassivos = rangePassivos.getValues();
  const richTextPassivos = rangePassivos.getRichTextValues(); // Necessário para pegar o link do emoji

  let htmlListaDespesas = "";
  let anexosArquivos = []; // Array para guardar os arquivos reais

  // Loop começa do 1 para pular cabeçalho
  for (let i = 1; i < dadosPassivos.length; i++) {
    const linha = dadosPassivos[i];
    const mesPassivo = linha[0];
    const anoPassivo = linha[1];
    const servico = linha[2];
    const valor = linha[3];

    // Verifica se é do mesmo Mês e Ano
    if (String(mesPassivo).toLowerCase() === String(mesAcerto).toLowerCase() && 
        String(anoPassivo) === String(anoAcerto)) {
      
      // 3.1. Adiciona à lista visual do e-mail
      let valFormatado = parseFloat(valor).toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});
      htmlListaDespesas += `<li style="margin-bottom: 5px;"><strong>${servico}:</strong> ${valFormatado}</li>`;

      // 3.2. EXTRAI O ANEXO DO DRIVE
      // A coluna do recibo é a índice 4 (Coluna E)
      const celulaRichText = richTextPassivos[i][4];
      const urlRecibo = celulaRichText ? celulaRichText.getLinkUrl() : null;

      if (urlRecibo && urlRecibo.includes("id=")) {
        const idDoc = urlRecibo.split("id=")[1];
        try {
          const arquivo = DriveApp.getFileById(idDoc);
          const blob = arquivo.getBlob();
          // Renomeia o arquivo para facilitar: "CEMIG - NomeOriginal.pdf"
          blob.setName(`${servico} - ${arquivo.getName()}`); 
          anexosArquivos.push(blob);
        } catch (e) {
          console.log(`Erro ao baixar anexo de ${servico}: ` + e.message);
        }
      }
    }
  }

  if (htmlListaDespesas === "") htmlListaDespesas = "<li><em>Sem detalhes lançados.</em></li>";

  // --- 4. PREPARAÇÃO DO QR CODE (IMAGEM INLINE) ---
  const cellQr = abaAcertos.getRange(ultimaLinha, 4);
  let idQrCode = "";
  
  // Lógica para pegar ID do QR Code (seja link ou fórmula)
  const rtQr = cellQr.getRichTextValue();
  const urlQr = rtQr ? rtQr.getLinkUrl() : "";
  const formulaQr = cellQr.getFormula();
  const matchFormula = formulaQr ? formulaQr.match(/id=([a-zA-Z0-9_-]+)/) : null;

  if (urlQr && urlQr.includes("id=")) idQrCode = urlQr.split("id=")[1];
  else if (matchFormula && matchFormula[1]) idQrCode = matchFormula[1];
  else if (typeof cellQr.getValue() === 'string' && cellQr.getValue().includes("id=")) idQrCode = cellQr.getValue().split("id=")[1];

  let blobQrCode = null;
  let imagensInline = {}; 
  if (idQrCode) {
    try {
      blobQrCode = DriveApp.getFileById(idQrCode).getBlob().setName("qrcode.png");
      imagensInline = { qrCodeImagem: blobQrCode }; 
    } catch (e) { console.log("Erro imagem QR Code"); }
  }
  const linkDiretoQr = idQrCode ? `https://drive.google.com/open?id=${idQrCode}` : "#";

  // --- 5. CONFIRMAÇÃO ---
  const resposta = ui.alert(
    `Confirmar Envio`,
    `Enviar para: ${listaDestinatarios.join(", ")}\n` +
    `Referência: ${mesAcerto}/${anoAcerto}\n` +
    `Total: ${valorFormatado}\n` +
    `Anexos encontrados: ${anexosArquivos.length} arquivos`,
    ui.ButtonSet.YES_NO
  );

  if (resposta !== ui.Button.YES) return;

  // --- 6. MONTAGEM E ENVIO ---
  const assunto = `🔔 Rateio Pampulha: ${mesAcerto}/${anoAcerto} (+ Comprovantes)`;

  const corpoEmail = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 500px;">
      <h2 style="color: #1155cc;">Rateio Disponível: ${mesAcerto}/${anoAcerto}</h2>
      
      <div style="background-color: #f9f9f9; padding: 15px; border-radius: 8px; border: 1px solid #ddd;">
        <p style="font-size: 18px; margin: 0;">Valor Individual:</p>
        <p style="font-size: 24px; font-weight: bold; color: #008000; margin: 5px 0;">${valorFormatado}</p>
      </div>

      <br>

      <div style="border: 1px solid #eee; padding: 10px; border-radius: 5px;">
        <p style="margin-top: 0;"><strong>📋 Composição dos Gastos:</strong></p>
        <ul style="padding-left: 20px; color: #555;">
          ${htmlListaDespesas}
        </ul>
        <p style="font-size: 12px; color: #888;">📎 <em>Os comprovantes originais estão anexados a este e-mail.</em></p>
      </div>

      <br>

      <p><strong>1. Pagamento via QR Code:</strong></p>
      <div style="text-align: center; margin: 10px 0;">
        ${blobQrCode ? 
          `<img src="cid:qrCodeImagem" style="width: 200px; height: 200px; border: 1px solid #ccc; padding: 5px;">` : 
          `<p style="color: #666;">(Imagem não carregada)</p>`
        }
        <br>
        <a href="${linkDiretoQr}" style="font-size: 14px; color: #1155cc;">🔗 Link alternativo do QR Code</a>
      </div>

      <p><strong>2. Pix Copia e Cola:</strong></p>
      <div style="background-color: #eee; padding: 10px; border-radius: 4px; font-family: monospace; font-size: 11px; word-break: break-all; color: #000;">
        ${chavePix || "Chave não detectada"}
      </div>

      <br><hr>
      <p style="font-size: 11px; color: #888;">Gestão Pampulha Automática</p>
    </div>
  `;

  let contagemSucesso = 0;
  listaDestinatarios.forEach(email => {
    try {
      MailApp.sendEmail({
        to: email,
        subject: assunto,
        htmlBody: corpoEmail,
        inlineImages: imagensInline,
        attachments: anexosArquivos // <--- AQUI VÃO OS COMPROVANTES
      });
      contagemSucesso++;
    } catch (e) {
      ui.alert(`❌ Erro ao enviar para ${email}: ${e.message}`);
    }
  });

  if (contagemSucesso > 0) {
    ui.alert(`✅ E-mail enviado com ${anexosArquivos.length} anexos!`);
  }
}