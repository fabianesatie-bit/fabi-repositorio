// =====================================================================
// SERVIÇO DE DISPARO DE NOTIFICAÇÕES (MALA DIRETA)
// =====================================================================

function dispararConvites() {
  try {
    // 1. Gera os dados cruzados chamando a nossa função principal do dash
    var processamento = carregarDadosDashboard();
    
    if (!processamento.sucesso) {
      return { sucesso: false, erro: "Falha ao ler a base de gerentes: " + processamento.erro };
    }
    
    var gerentes = processamento.payload.gerentes;
    var urlApp = ScriptApp.getService().getUrl(); // Pega a URL exata deste formulário
    var contadorEnviados = 0;
    
    // 2. Itera apenas sobre os pendentes e faz o disparo
    for (var i = 0; i < gerentes.length; i++) {
      if (gerentes[i].status === "Pendente" && gerentes[i].email) {
        
        var emailDestino = gerentes[i].email;
        var assunto = "Inscrição para Posicionamento estrategico";
        var corpoHtml = `
          <div style="font-family: Arial, sans-serif; color: #334155; padding: 20px; max-width: 600px; margin: 0 auto; border: 1px solid #E2E8F0; border-radius: 10px;">
            <h2 style="color: #6B21A8;">Olá, ${gerentes[i].nome}!</h2>
            <p style="font-size: 16px; line-height: 1.5;">Convidamos você a preencher sua inscrição para o Posicionamento Estratégico.</p>
            <p style="font-size: 16px; line-height: 1.5;">Acesse o link seguro abaixo para registrar suas informações de logística e transporte:</p>
            <br>
            <div style="text-align: center; margin: 30px 0;">
              <a href="${urlApp}" style="background-color: #7C3AED; color: #FFFFFF; padding: 14px 28px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block;">
                Acessar Formulário de Inscrição
              </a>
            </div>
            <br>
            <p style="font-size: 16px; line-height: 1.5; color: #64748B;"><strong>Enviaremos maiores informações posteriormente.</strong></p>
            <hr style="border: 0; border-top: 1px solid #E2E8F0; margin: 20px 0;">
            <p style="font-size: 12px; color: #94A3B8; text-align: center;">Equipe de Organização - Evento Regional<br>Magazine Luiza</p>
          </div>
        `;
        
        // Dispara utilizando o GmailApp
        GmailApp.sendEmail(emailDestino, assunto, "Convidamos você a preencher sua inscrição. Acesse: " + urlApp, {
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

/**
 * FUNÇÃO DE TESTE: Envia um e-mail de demonstração apenas para o desenvolvedor logado.
 * (Usado para validar layout e responsividade do e-mail antes do disparo real).
 */
function testarLayoutEmail() {
  var emailDestino = Session.getActiveUser().getEmail(); 
  
  var nomeTeste = "Fabiane (Demonstração)";
  var urlApp = ScriptApp.getService().getUrl(); 
  if (!urlApp || urlApp === "") {
    urlApp = "https://script.google.com/macros/s/LINK_DE_TESTE/exec";
  }
  
  var assunto = "[TESTE DE LAYOUT] Inscrição para Posicionamento estrategico";
  
  var corpoHtml = `
    <div style="font-family: Arial, sans-serif; color: #334155; padding: 20px; max-width: 600px; margin: 0 auto; border: 1px solid #E2E8F0; border-radius: 10px;">
      <h2 style="color: #6B21A8;">Olá, ${nomeTeste}!</h2>
      <p style="font-size: 16px; line-height: 1.5;">Convidamos você a preencher sua inscrição para o Posicionamento Estratégico.</p>
      <p style="font-size: 16px; line-height: 1.5;">Acesse o link seguro abaixo para registrar suas informações de logística e transporte:</p>
      <br>
      <div style="text-align: center; margin: 30px 0;">
        <a href="${urlApp}" style="background-color: #7C3AED; color: #FFFFFF; padding: 14px 28px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block;">
          Acessar Formulário de Inscrição
        </a>
      </div>
      <br>
      <p style="font-size: 16px; line-height: 1.5; color: #64748B;"><strong>Enviaremos maiores informações posteriormente.</strong></p>
      <hr style="border: 0; border-top: 1px solid #E2E8F0; margin: 20px 0;">
      <p style="font-size: 12px; color: #94A3B8; text-align: center;">Equipe de Organização - Evento Regional<br>Magazine Luiza</p>
    </div>
  `;
  
  GmailApp.sendEmail(emailDestino, assunto, "Acesse o link para preencher sua inscrição.", {
    htmlBody: corpoHtml
  });
  
  Logger.log("Sucesso! E-mail de demonstração enviado para: " + emailDestino);
}
