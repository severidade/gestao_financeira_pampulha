function notificar_novo_acerto() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- CONFIGURAÇÃO ---
  const listaDestinatarios = ["oliveira.severo@gmail.com"];
  
  const abaAcertos = ss.getSheetByName("🤝 Acertos_Mensais_Dados_Brutos");
  const abaPassivos = ss.getSheetByName("💸 Passivos_Dados_Brutos");

  if (!abaAcertos || !abaPassivos) { ui.alert("Abas não encontradas!"); return; }

  const ultimaLinha = abaAcertos.getLastRow();
  if (ultimaLinha < 2) { ui.alert("Tabela vazia."); return; }

  // --- FUNÇÃO PARA LIMPAR ACENTOS (O SEGREDO DO "MARCO" vs "MARÇO") ---
  function normalizar(texto) {
    return String(texto)
      .toLowerCase()
      .trim()
      .normalize("NFD") // Separa os acentos das letras
      .replace(/[\u0300-\u036f]/g, ""); // Remove os acentos
  }

  // Dados do Acerto
  const mesAcerto = normalizar(abaAcertos.getRange(ultimaLinha, 1).getValue()); // Normaliza aqui
  const anoAcerto = String(abaAcertos.getRange(ultimaLinha, 2).getValue()).trim();
  const valorFormatado = abaAcertos.getRange(ultimaLinha, 3).getDisplayValue();
  const chavePix = abaAcertos.getRange(ultimaLinha, 5).getValue();

  // --- BUSCA NA TABELA PASSIVOS ---
  const dados = abaPassivos.getDataRange().getValues();
  const dadosRicos = abaPassivos.getDataRange().getRichTextValues();

  let htmlLista = "";
  let anexos = [];

  // Apenas para ver no log se funcionou
  const mesOriginal = abaAcertos.getRange(ultimaLinha, 1).getValue();
  console.log(`Buscando: "${mesOriginal}" (lido como: "${mesAcerto}") / ${anoAcerto}`);

  // Loop começa do 1 para pular cabeçalho
  for (let i = 1; i < dados.length; i++) {
    // Índices: Col A[0]=Mês | Col B[1]=Ano | Col C[2]=Serviço | Col D[3]=Valor | Col E[4]=RECIBO
    
    // Normaliza o mês da tabela Passivos também (Marco vira marco)
    let mesPassivo = normalizar(dados[i][0]);
    let anoPassivo = String(dados[i][1]).trim();
    
    // Agora a comparação funciona mesmo se for "Março" vs "Marco"
    if (mesPassivo === mesAcerto && anoPassivo === anoAcerto) {
      
      let servico = dados[i][2]; 
      let valor = parseFloat(dados[i][3]).toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});
      
      htmlLista += `<li><strong>${servico}:</strong> ${valor}</li>`;

      // --- EXTRAÇÃO DO LINK DA COLUNA E (ÍNDICE 4) ---
      let celulaRica = dadosRicos[i][4]; 
      let url = celulaRica ? celulaRica.getLinkUrl() : null;

      // Fallback (se não for RichText, tenta texto puro)
      if (!url && String(dados[i][4]).includes("http")) {
        url = dados[i][4];
      }

      if (url) {
        try {
          // Limpeza do ID
          let idArquivo = "";
          if (url.includes("id=")) idArquivo = url.split("id=")[1].split("&")[0];
          else if (url.includes("/d/")) idArquivo = url.split("/d/")[1].split("/")[0];

          if (idArquivo) {
            let arquivo = DriveApp.getFileById(idArquivo);
            let blob = arquivo.getBlob();
            blob.setName(`${servico} - Comprovante.pdf`);
            anexos.push(blob);
            console.log(`✅ Anexo encontrado: ${servico}`);
          }
        } catch (e) {
          console.log(`❌ Erro anexo ${servico}: ` + e.message);
        }
      }
    }
  }

  // --- E-MAIL ---
  if (htmlLista === "") htmlLista = "<li>Nenhum detalhe encontrado.</li>";
  
  // QR Code (Aba Acertos, Coluna D -> Índice 4 no getRange)
  let qrCodeBlob = null;
  let qrCodeInline = {};
  try {
    let cellQr = abaAcertos.getRange(ultimaLinha, 4);
    let qrUrl = cellQr.getRichTextValue() ? cellQr.getRichTextValue().getLinkUrl() : null;
    if (!qrUrl && String(cellQr.getValue()).includes("http")) qrUrl = cellQr.getValue();

    if (qrUrl) {
      let idQr = "";
      if (qrUrl.includes("id=")) idQr = qrUrl.split("id=")[1].split("&")[0];
      else if (qrUrl.includes("/d/")) idQr = qrUrl.split("/d/")[1].split("/")[0];
      
      if (idQr) {
        qrCodeBlob = DriveApp.getFileById(idQr).getBlob().setName("qrcode.png");
        qrCodeInline = { qrImagem: qrCodeBlob };
      }
    }
  } catch(e) { console.log("Sem QR Code"); }

  // Confirmação
  const confirma = ui.alert(`Confirmar Envio`, 
    `Mês: ${mesAcerto}/${anoAcerto}\nAnexos encontrados: ${anexos.length}`, 
    ui.ButtonSet.YES_NO);
  
  if (confirma !== ui.Button.YES) return;

  const htmlBody = `
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f6f6; padding:20px 0;">
    <tr>
      <td align="center">

        <!-- CONTAINER -->
        <table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:8px; font-family:Arial,sans-serif; color:#333;">
          
          <!-- CABEÇALHO -->
          <tr>
            <td style="padding:24px; text-align:center;">
              <h2 style="margin:0; color:#00008B;">
                Contas da Pampulha · ${mesOriginal}/${anoAcerto}
              </h2>
            </td>
          </tr>

          <!-- VALOR -->
          <tr>
            <td style="padding:16px 24px;">
              <div style="background:#f1f1f1; padding:14px; border-radius:6px; text-align:center;">
                <span style="font-size:16px;">
                  Valor da cota individual:
                  <strong>${valorFormatado}</strong>
                </span>
              </div>
            </td>
          </tr>

          <!-- DETALHAMENTO -->
          <tr>
            <td style="padding:16px 24px;">
              <h4 style="margin:0 0 8px 0;">Detalhamento</h4>
              <ul style="margin:0; padding-left:20px;">
                ${htmlLista}
              </ul>
              <p style="font-size:12px; color:#666; margin-top:12px;">
                📎 <em>Os comprovantes seguem anexos.</em>
              </p>
            </td>
          </tr>

          <!-- BLOCO PIX -->
          <tr>
            <td style="background:#f5f6f6; padding:32px 24px; text-align:center;">

              <p style="font-size:16px; font-weight:700; margin:0 0 8px 0; color:#041e18;">
                Pix QR Code
              </p>
              <p style="font-size:14px; margin:0 0 24px 0; color:#041e18;">
                Leia o QR Code usando seu aplicativo de pagamento.
              </p>

              ${qrCodeBlob ? `
                <img src="cid:qrImagem"
                    width="180"
                    alt="QR Code Pix"
                    style="display:block; margin:0 auto 24px auto; border:1px solid #ccc;">
              ` : ''}

              <p style="font-size:16px; font-weight:700; margin:0 0 8px 0; color:#041e18;">
                Pix Copia e Cola
              </p>
              <p style="font-size:14px; margin:0 0 16px 0; color:#041e18;">
                No aplicativo do seu banco, selecione Pix Copia e Cola e insira o código abaixo.
              </p>

              <div style="background:#ffffff; border:1px solid #b1b9b7; border-radius:8px; padding:16px; text-align:left;">
                <p style="margin:0; font-size:14px; line-height:20px; word-break:break-all; color:#041e18;">
                  ${chavePix}
                </p>
              </div>

            </td>
          </tr>

        </table>

      </td>
    </tr>
  </table>
  `;


  listaDestinatarios.forEach(email => {
    try {
      MailApp.sendEmail({
        to: email,
        subject: `[ 💸 Pampulha ] | Contas de ${mesOriginal}/${anoAcerto}`,
        htmlBody: htmlBody,
        inlineImages: qrCodeInline,
        attachments: anexos
      });
    } catch (e) {
      console.log("Erro envio: " + e.message);
    }
  });

  ui.alert(`✅ E-mail enviado com ${anexos.length} anexo(s)!`);
}