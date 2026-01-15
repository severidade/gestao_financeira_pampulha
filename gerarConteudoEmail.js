// Alteração 1: Adicionei o parâmetro 'mensagemPersonalizada'
function gerarConteudoEmail(chaveSelecionada, mensagemPersonalizada) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAcertos = ss.getSheetByName("🤝 Acertos_Mensais_Dados_Brutos");
  const abaPassivos = ss.getSheetByName("💸 Passivos_Dados_Brutos");
  
  // CONSTANTES
  const NUMERO_PESSOAS = 3; 
  const [mesAlvo, anoAlvo] = chaveSelecionada.split("|");

  // --- FUNÇÃO AUXILIAR PARA LIMPAR NÚMEROS ---
  const converterParaFloat = (v) => {
    if (typeof v === 'number') return v;
    if (!v) return 0;
    return parseFloat(String(v).replace("R$", "").replace(/\./g, "").replace(",", ".").trim()) || 0;
  };

  // --- 1. LÓGICA DE TEXTO AMIGÁVEL ---
  const regexExtra = /^(.*?)\s*\((\d+)\)$/;
  const match = mesAlvo.match(regexExtra);

  let tituloFormatado = "";
  let nomeMesLimpo = "";

  if (match) {
    nomeMesLimpo = match[1].charAt(0).toUpperCase() + match[1].slice(1); 
    let numeroExtra = match[2];
    tituloFormatado = `Gastos Complementares ${numeroExtra} - ${nomeMesLimpo}`;
  } else {
    nomeMesLimpo = mesAlvo.charAt(0).toUpperCase() + mesAlvo.slice(1);
    tituloFormatado = `Despesas de ${nomeMesLimpo}`;
  }

  // --- 2. BUSCAR DADOS DO ACERTO ---
  const dadosAcertos = abaAcertos.getDataRange().getValues();
  const ricosAcertos = abaAcertos.getDataRange().getRichTextValues();
  let dadosCobranca = null;

  for (let i = 1; i < dadosAcertos.length; i++) {
    if (String(dadosAcertos[i][0]) === mesAlvo && String(dadosAcertos[i][1]) === anoAlvo) {
      let dataVenc = dadosAcertos[i][2];
      let textoVencimento = dataVenc instanceof Date ? Utilities.formatDate(dataVenc, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : String(dataVenc);

      dadosCobranca = {
        titulo: tituloFormatado, 
        ano: anoAlvo,
        vencimento: textoVencimento,
        valor: dadosAcertos[i][3], 
        chavePix: dadosAcertos[i][5],
        qrRichText: ricosAcertos[i][4],
        qrTextoBackup: dadosAcertos[i][4]
      };
      break; 
    }
  }

  if (!dadosCobranca) throw new Error("Cobrança não encontrada na tabela de Acertos.");

  let valorCobradoNumerico = converterParaFloat(dadosCobranca.valor);
  let valorFinalDisplay = valorCobradoNumerico.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});

  // --- 3. BUSCAR PASSIVOS E SOMAR ---
  const dadosPassivos = abaPassivos.getDataRange().getValues();
  const ricosPassivos = abaPassivos.getDataRange().getRichTextValues();
  
  let htmlLista = "";
  let anexos = [];
  let totalPassivosAcumulado = 0; 
  
  for (let i = 1; i < dadosPassivos.length; i++) {
    let passivoMes = String(dadosPassivos[i][0]).trim();
    let passivoAno = String(dadosPassivos[i][1]).trim();

    if (passivoMes === mesAlvo && passivoAno === anoAlvo) {
      let servico = dadosPassivos[i][2];
      let valPassivoRaw = dadosPassivos[i][3];
      
      let valPassivoNum = converterParaFloat(valPassivoRaw);
      totalPassivosAcumulado += valPassivoNum; 

      let valPassivoFmt = valPassivoNum.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});
      htmlLista += `<li><strong>${servico}:</strong> ${valPassivoFmt}</li>`;

      // Anexos
      let celulaRica = ricosPassivos[i][4]; 
      let url = celulaRica ? celulaRica.getLinkUrl() : null;
      if (!url && String(dadosPassivos[i][4]).includes("http")) url = dadosPassivos[i][4];

      if (url) {
        try {
          let idArquivo = "";
          if (url.includes("id=")) idArquivo = url.split("id=")[1].split("&")[0];
          else if (url.includes("/d/")) idArquivo = url.split("/d/")[1].split("/")[0];
          
          if (idArquivo) {
            let arquivo = DriveApp.getFileById(idArquivo);
            let nomeOriginal = arquivo.getName();
            let extensao = ".pdf"; 

            if (nomeOriginal.indexOf(".") !== -1) {
              extensao = nomeOriginal.substring(nomeOriginal.lastIndexOf("."));
            }

            let blob = arquivo.getBlob().setName(`${servico}${extensao}`);
            anexos.push(blob);
          }
        } catch (e) { console.log("Erro anexo: " + e.message); }
      }
    }
  }
  
  if (htmlLista === "") htmlLista = "<li>Nenhum detalhe encontrado.</li>";

  let valorTotalDisplay = totalPassivosAcumulado.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'});

  // --- 4. QR CODE BLOB ---
  let qrCodeBlob = null;
  try {
    let qrUrl = dadosCobranca.qrRichText ? dadosCobranca.qrRichText.getLinkUrl() : null;
    if (!qrUrl && String(dadosCobranca.qrTextoBackup).includes("http")) qrUrl = dadosCobranca.qrTextoBackup;
    
    if (qrUrl) {
      let idQr = "";
      if (qrUrl.includes("id=")) idQr = qrUrl.split("id=")[1].split("&")[0];
      else if (qrUrl.includes("/d/")) idQr = qrUrl.split("/d/")[1].split("/")[0];
      if (idQr) qrCodeBlob = DriveApp.getFileById(idQr).getBlob().setName("qrcode.png");
    }
  } catch(e) {}

  // --- 5. PREPARAR MENSAGEM CUSTOMIZADA (LIMPA) ---
  let htmlMensagemExtra = "";
  if (mensagemPersonalizada && mensagemPersonalizada.trim() !== "") {
    // Substitui quebra de linha (\n) por <br>
    let msgFormatada = mensagemPersonalizada.replace(/\n/g, '<br>');
    
    // Agora é um <p> simples, sem caixa amarela
    htmlMensagemExtra = `
      <tr>
        <td style="padding: 0 24px 15px 24px; text-align:left;">
          <p style="font-size: 14px; color: #333; line-height: 1.6; margin: 0;">
            ${msgFormatada}
          </p>
        </td>
      </tr>
    `;
  }

  // --- 6. MONTAR HTML ---
  const htmlBody = `
    <table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f6f6; padding:20px 0;">
      <tr>
        <td align="center">
          <table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff; border-radius:8px; font-family:Arial,sans-serif; color:#333;">
            
            <tr>
              <td style="padding:24px; text-align:center;">
                <h2 style="margin:0; color:#00008B;">Contas da Pampulha</h2>
                <p style="font-size:16px; color:#555; margin-top:8px; font-weight:bold;">
                   ${dadosCobranca.titulo} / ${dadosCobranca.ano}
                </p>
              </td>
            </tr>

            ${htmlMensagemExtra}
            
            <tr>
              <td style="padding:16px 24px;">
                <div style="background:#e8f0fe; padding:20px; border-radius:8px; text-align:center; border: 1px solid #d2e3fc;">
                  
                  <div style="font-size:14px; color:#1155cc; margin-bottom:10px; font-weight:bold; text-transform:uppercase;">
                    Valor da cota individual com vencimento: ${dadosCobranca.vencimento}
                  </div>
                  
                  <div style="font-size:28px; font-weight:bold; color:#000;">${valorFinalDisplay}</div>
                  
                  <div style="font-size:12px; color:#666; margin-top:5px;">
                    (${valorTotalDisplay} dividido por ${NUMERO_PESSOAS})
                  </div>

                </div>
              </td>
            </tr>
            <tr>
              <td style="padding:16px 24px;">
                <h4 style="margin:0 0 8px 0; border-bottom:1px solid #eee; padding-bottom:5px;">Detalhamento</h4>
                <ul style="margin:0; padding-left:20px; line-height: 1.6;">${htmlLista}</ul>
                <p style="font-size:12px; color:#666; margin-top:15px;">📎 <em>${anexos.length} comprovante(s) em anexo.</em></p>
              </td>
            </tr>
            <tr>
              <td style="background:#f5f6f6; padding:32px 24px; text-align:center; border-bottom-left-radius:8px; border-bottom-right-radius:8px;">
                ${qrCodeBlob ? `<p style="font-weight:700;">Pix QR Code</p><img src="cid:qrImagem" width="160" style="display:block; margin:0 auto 20px auto; border:4px solid white;">` : ''}
                <p style="font-weight:700;">Pix Copia e Cola</p>
                <div style="background:#ffffff; border:1px solid #ccc; border-radius:6px; padding:12px; text-align:left; word-break:break-all; font-family:monospace; font-size:12px;">${dadosCobranca.chavePix}</div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  `;

  return {
    htmlBody: htmlBody,
    assunto: `[ 💸 Pampulha ] ${dadosCobranca.titulo} / ${dadosCobranca.ano}`,
    anexos: anexos,
    qrBlob: qrCodeBlob
  };
}

