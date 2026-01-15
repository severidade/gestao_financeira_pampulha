function abrirPreviewEmail(chaveSelecionada) {
  
  // 1. Gera os dados iniciais (sem mensagem extra)
  const dados = gerarConteudoEmail(chaveSelecionada, "");

  // 2. Prepara imagem QR Code para visualização (Base64)
  let corpoEmailPreview = dados.htmlBody.replace('src="cid:qrImagem"', `src="data:image/png;base64,${dados.qrBlob ? Utilities.base64Encode(dados.qrBlob.getBytes()) : ''}"`);

  const htmlModal = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { margin: 0; padding: 0; background: #eee; font-family: sans-serif; display: flex; flex-direction: column; height: 100vh; }
          
          /* BARRA DE FERRAMENTAS (FIXA NO TOPO) */
          .toolbar { 
            background: #333; color: white; padding: 10px 20px; 
            display: flex; justify-content: space-between; align-items: center;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2); flex-shrink: 0;
          }
          
          /* ÁREA DE INPUT MENSAGEM (FIXA ABAIXO DA TOOLBAR) */
          .input-area {
            background: #e0e0e0; padding: 15px 20px; border-bottom: 1px solid #ccc;
            display: flex; gap: 10px; align-items: flex-start; flex-shrink: 0;
          }
          
          textarea {
            flex-grow: 1; height: 40px; padding: 8px; border-radius: 4px; border: 1px solid #999;
            font-family: sans-serif; resize: vertical; font-size: 13px;
          }
          
          .btn-refresh {
            background: #673AB7; color: white; border: none; padding: 0 15px; height: 58px; 
            border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 12px;
          }
          .btn-refresh:hover { background: #5e35b1; }

          /* ÁREA DE PREVIEW (COM SCROLL) */
          .preview-container {
             flex-grow: 1; overflow-y: auto; padding: 20px;
          }
          .preview-content { max-width: 650px; margin: 0 auto; background: white; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
          
          /* BOTÕES GERAIS */
          .btn { padding: 10px 20px; border-radius: 4px; border: none; font-weight: bold; cursor: pointer; font-size: 14px; }
          .btn-cancel { background: transparent; border: 1px solid #777; color: #ccc; margin-right: 10px; }
          .btn-back { background: #1155cc; color: white; margin-right: 10px; }
          .btn-send { background: #28a745; color: white; }
          .btn-send:hover { background: #218838; }
          .info { font-size: 12px; color: #ccc; display: flex; align-items: center; gap: 10px;}
        </style>
      </head>
      <body>
        
        <div class="toolbar">
          <div class="info">
             <button class="btn btn-cancel" onclick="google.script.host.close()">❌ Fechar</button>
             <span><strong>Anexos:</strong> ${dados.anexos.length}</span>
          </div>
          <div>
            <button class="btn btn-back" onclick="voltarParaSeletor()">⬅️ Voltar</button>
            <button class="btn btn-send" id="btnEnviar" onclick="confirmarEnvio()">✅ Confirmar e Enviar</button>
          </div>
        </div>

        <div class="input-area">
          <textarea id="msgExtra" placeholder="Digite uma mensagem ou aviso opcional para aparecer no e-mail..."></textarea>
          <button class="btn-refresh" onclick="atualizarVisualizacao()">🔄 Atualizar<br>Visualização</button>
        </div>
        
        <div class="preview-container">
          <div class="preview-content" id="conteudoEmail">
            ${corpoEmailPreview}
          </div>
        </div>

        <script>
          // 1. Atualiza o HTML do preview sem fechar a janela
          function atualizarVisualizacao() {
            var texto = document.getElementById('msgExtra').value;
            var divPreview = document.getElementById('conteudoEmail');
            
            // Coloca uma opacidade para indicar carregamento
            divPreview.style.opacity = "0.5";
            
            google.script.run
              .withSuccessHandler(function(novoHtml) {
                 divPreview.innerHTML = novoHtml;
                 divPreview.style.opacity = "1";
              })
              .recuperarHtmlAtualizado("${chaveSelecionada}", texto);
          }

          // 2. Envia o e-mail com a mensagem
          function confirmarEnvio() {
            var btn = document.getElementById('btnEnviar');
            var texto = document.getElementById('msgExtra').value; // Pega o texto
            
            btn.innerHTML = "⏳ Enviando...";
            btn.disabled = true;
            
            google.script.run
              .withSuccessHandler(function() { 
                 alert('E-mail enviado com sucesso!'); 
                 google.script.host.close(); 
              })
              .withFailureHandler(function(e) {
                 alert('Erro: ' + e);
                 btn.disabled = false;
                 btn.innerHTML = "Tentar Novamente";
              })
              // Passa a chave E a mensagem para o envio
              .enviarCobrancaEmail("${chaveSelecionada}", texto);
          }

          function voltarParaSelecao() {
            google.script.run.abrirPainelSelecaoEmail();
          }
        </script>
      </body>
    </html>
  `;

  const output = HtmlService.createHtmlOutput(htmlModal).setWidth(900).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(output, `Pré-visualização: ${dados.assunto}`);
}