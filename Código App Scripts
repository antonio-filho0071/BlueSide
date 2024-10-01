function gerarEEnviarIngressos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  var templateId = 'ID-GOOGLE-SLIDES'; // Substitua pelo ID do seu modelo no Google Slides
  var folder = DriveApp.getFolderById('ID-PASTA-GOOGLEDRIVE'); // Substitua pelo ID da pasta onde deseja salvar os PDFs
  
  for (var i = 1; i < data.length; i++) { // Começa em 1 para pular o cabeçalho
    var nome = data[i][X]; // Ajuste o índice X conforme a coluna do nome
    var email = data[i][X]; // Ajuste o índice X conforme a coluna do e-mail
    var numeroInscricao = data[i][X]; // Ajuste o índice X conforme a coluna do número de inscrição
    
    // Copiar o modelo
    var copia = DriveApp.getFileById(templateId).makeCopy('Ingresso - ' + nome);
    var copiaId = copia.getId();
    
    // Abrir a apresentação
    var presentation = SlidesApp.openById(copiaId);
    var slides = presentation.getSlides();
    
    // Substituir textos no slide
    slides.forEach(function(slide) {
      slide.replaceAllText('{{NOME}}', nome); // Substitua {{NOME}} pelo texto correspondente no seu modelo
      slide.replaceAllText('{{NUMERO_INSCRICAO}}', numeroInscricao); // Substitua {{NUMERO_INSCRICAO}} pelo texto correspondente no seu modelo
    });
    
    // Salvar como PDF
    presentation.saveAndClose();
    var pdf = DriveApp.getFileById(copiaId).getAs(MimeType.PDF);
    var pdfFile = folder.createFile(pdf);
    
    // Envio do e-mail com anexo
    var assunto = "Seu Ingresso para o Evento [Nome do Evento]";
    var mensagem = "Olá " + nome + ",\n\n" +
                   "Obrigado por se inscrever no [Nome do Evento]!\n\n" +
                   "Anexo está o seu ingresso com o número de inscrição: " + numeroInscricao + ".\n\n" +
                   "Estamos ansiosos para vê-lo!\n\n" +
                   "Atenciosamente,\n" +
                   "[Seu Nome]";
    
    MailApp.sendEmail({
      to: email,
      subject: assunto,
      body: mensagem,
      attachments: [pdfFile]
    });
    
    // Opcional: Excluir a cópia do Google Slides
    DriveApp.getFileById(copiaId).setTrashed(true);
  }
}

