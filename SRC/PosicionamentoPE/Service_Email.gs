// =====================================================================
// SERVIÇO DE DISPARO DE NOTIFICAÇÕES (MALA DIRETA) - ECOSSISTEMA GP360
// =====================================================================

function dispararConvites() {
  try {
    var processamento = carregarDadosDashboard();
    
    if (!processamento.sucesso) {
      return { sucesso: false, erro: "Falha ao ler a base de gerentes: " + processamento.erro };
    }
    
    var gerentes = processamento.payload.gerentes;
    var urlApp = ScriptApp.getService().getUrl(); 
    var contadorEnviados = 0;
    
    // LISTA DE DIRETORIA MG/CO PERMITIDA PARA RECEBIMENTO DE E-MAILS
    var regionaisPermitidas = ["BRASILIA", "UBERLANDIA", "JUIZ DE FORA", "BELO HORIZONTE", "CAMPO GRANDE", "CUIABA"];
    
    for (var i = 0; i < gerentes.length; i++) {
      
      var regionalNormalizada = String(gerentes[i].regional || "").toUpperCase();
      
      // TRAVA DE FILTRO MG/CO E ESCRITÓRIO:
      // O e-mail só será enviado se a regional da pessoa existir dentro do array acima.
      var permiteEnvio = false;
      for (var r = 0; r < regionaisPermitidas.length; r++) {
         if (regionalNormalizada.indexOf(regionaisPermitidas[r]) !== -1) {
            permiteEnvio = true;
            break;
         }
      }

      // Se não pertencer à lista (ex: ESCRITORIO, SP, NE, etc.), o robô ignora e pula para o próximo
      if (!permiteEnvio) {
         continue; 
      }

      if (gerentes[i].status === "Pendente" && gerentes[i].email) {
        
        var emailDestino = gerentes[i].email;
        var assunto = "Convite Oficial: Posicionamento Estratégico Outubro 2026 &#10024;";
        
        var corpoHtml = `
          <div style="font-family: 'Segoe UI', Tahoma, Arial, sans-serif; color: #334155; padding: 20px; max-width: 600px; margin: 0 auto; border: 1px solid #E2E8F0; border-radius: 12px; background-color: #F8FAFC;">
            
            <div style="text-align: center; margin-bottom: 25px;">
               <img src="https://drive.google.com/uc?export=view&id=1dD9Z-1KK2iPVAsLorbljK9H6lg3GIHc3" alt="Posicionamento Outubro 2026" style="max-width: 100%; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.05);">
            </div>

            <h2 style="color: #6B21A8; text-align: center; margin-top: 0;">Olá, ${gerentes[i].nome}! &#10024;</h2>
            
            <div style="background-color: #FFFBEB; border: 1px solid #FDE68A; padding: 15px; border-radius: 8px; text-align: center; margin-bottom: 25px;">
               <p style="margin: 0; font-size: 18px; color: #92400E;">&#128197; <strong>Reserve essa data: 26 e 27/10!</strong></p>
            </div>

            <p style="font-size: 16px; line-height: 1.6; text-align: center;">Faça sua inscrição conforme o local a ser realizado o evento da sua regional.</p>
            <br>

            <div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-left: 4px solid #2563EB; border-radius: 8px; padding: 15px; margin-bottom: 12px;">
               <p style="margin: 0 0 5px 0; color: #1E40AF; font-size: 14px; font-weight: bold;">&#128205; LOCAL DO EVENTO: Regional Campo Grande</p>
               <p style="margin: 0; font-size: 15px; color: #475569;">&#127976; <strong>Hotel Deville</strong><br><span style="font-size: 13px;">End: Av. Mato Grosso, 4250 - Carandá Bosque, Campo Grande - MS, 79002-232</span></p>
            </div>

            <div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-left: 4px solid #10B981; border-radius: 8px; padding: 15px; margin-bottom: 12px;">
               <p style="margin: 0 0 5px 0; color: #065F46; font-size: 14px; font-weight: bold;">&#128205; LOCAL DO EVENTO: Regional Cuiabá</p>
               <p style="margin: 0; font-size: 15px; color: #475569;">&#127976; <strong>Hotel Paiaguas</strong><br><span style="font-size: 13px;">End: Av. Historiador Rubens de Mendonça, 1718 - Bosque da Saúde, Cuiabá - MT, 78050-030</span></p>
            </div>

            <div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-left: 4px solid #F59E0B; border-radius: 8px; padding: 15px; margin-bottom: 12px;">
               <p style="margin: 0 0 5px 0; color: #92400E; font-size: 14px; font-weight: bold;">&#128205; LOCAL DO EVENTO: Regionais Brasília e Uberlândia</p>
               <p style="margin: 0; font-size: 15px; color: #475569;">&#127976; <strong>Hotel San Marino</strong><br><span style="font-size: 13px;">End: R. 5, 1090 - St. Oeste, Goiânia - GO, 74115-060</span></p>
            </div>

            <div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-left: 4px solid #F43F5E; border-radius: 8px; padding: 15px; margin-bottom: 25px;">
               <p style="margin: 0 0 5px 0; color: #9F1239; font-size: 14px; font-weight: bold;">&#128205; LOCAL DO EVENTO: Regionais Belo Horizonte e Juiz de Fora</p>
               <p style="margin: 0; font-size: 15px; color: #475569;">&#127976; <strong>Hotel San Diego Suítes Lourdes</strong><br><span style="font-size: 13px;">End: Av. Álvares Cabral, 1181 - Lourdes, Belo Horizonte - MG, 30170-001</span></p>
            </div>

            <div style="text-align: center; margin: 35px 0;">
              <a href="${urlApp}" style="background-color: #7C3AED; color: #FFFFFF; padding: 16px 32px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block; box-shadow: 0 4px 6px -1px rgba(124, 58, 237, 0.3);">
                &#9989; Acessar Formulário de Inscrição
              </a>
            </div>
            
            <hr style="border: 0; border-top: 1px solid #E2E8F0; margin: 30px 0 20px 0;">
            <p style="font-size: 12px; color: #94A3B8; text-align: center;">Equipe de Organização - Evento Regional<br>Magazine Luiza</p>
          </div>
        `;
        
        GmailApp.sendEmail(emailDestino, assunto, "Convidamos você a preencher sua inscrição para o Posicionamento Estratégico. Acesse: " + urlApp, {
          htmlBody: corpoHtml
        });
        
        contadorEnviados++;
      }
    }
    
    return { sucesso: true, enviados: contadorEnviados };
    
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

function testarLayoutEmail() {
  // Mantida inalterada da versão anterior (mesmo layout)
  var emailDestino = Session.getActiveUser().getEmail(); 
  var nomeTeste = "Fabiane (Demonstração)";
  var urlApp = ScriptApp.getService().getUrl(); 
  if (!urlApp || urlApp === "") urlApp = "https://script.google.com/macros/s/LINK_DE_TESTE/exec";
  
  var assunto = "[TESTE DE LAYOUT] Convite Oficial: Posicionamento Estratégico Outubro 2026 &#10024;";
  
  var corpoHtml = `
    <!-- ... Todo o html idêntico ao de cima que já testamos ... -->
    <div style="font-family: 'Segoe UI', Tahoma, Arial, sans-serif; color: #334155; padding: 20px; max-width: 600px; margin: 0 auto; border: 1px solid #E2E8F0; border-radius: 12px; background-color: #F8FAFC;">
      <div style="text-align: center; margin-bottom: 25px;">
         <img src="https://drive.google.com/uc?export=view&id=1dD9Z-1KK2iPVAsLorbljK9H6lg3GIHc3" alt="Posicionamento Outubro 2026" style="max-width: 100%; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.05);">
      </div>
      <h2 style="color: #6B21A8; text-align: center; margin-top: 0;">Olá, ${nomeTeste}! &#10024;</h2>
      <div style="text-align: center; margin: 35px 0;">
        <a href="${urlApp}" style="background-color: #7C3AED; color: #FFFFFF; padding: 16px 32px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block;">
          &#9989; Acessar Formulário de Inscrição
        </a>
      </div>
    </div>
  `;
  
  GmailApp.sendEmail(emailDestino, assunto, "Acesse o link.", { htmlBody: corpoHtml });
  Logger.log("Sucesso! E-mail de demonstração enviado.");
}
