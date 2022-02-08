/**
 * Adiciona a secao de feedback para todos os forms em uma pasta
 */
function addFeedback2All(){
  var info = getInfo();
  var folderID = getFolderId(info.pasta_destino);
  var files = DriveApp.getFolderById(folderID).getFiles();
  var pessoa = 0;

  // Percorre todos os arquivos da pasta
  while (files.hasNext()) {
    pessoa += 1;
    var file = files.next();
    Logger.log(file.getName());

    addFeedback(file.getId());
  }
}

/**
 * Adiciona uma seção com feedbacks para cada área avaliada em um formulário
 */
function addFeedback(id){
  var form = FormApp.openById(id);

  // Adiciona o titulo da secao de feedback
  form.addSectionHeaderItem()
      .setTitle('Feedback')
      .setHelpText('Deixe aqui sugestões, comentários ou elogios relacionados a cada um dos seguintes aspectos');

  // Adiciona pergunta realizacao da expectativa geral
  form.addParagraphTextItem()
  .setTitle('Expectativas que você tinha ao entrar na Inderios')
  .setValidation(null);
   
  // Adiciona pergunta de feedback das macroareas
  form.addParagraphTextItem()
  .setTitle('Desenvolvimento nas macroareas')
  .setValidation(null);

  // Adiciona pergunta de feedback  nas soft-skills
  form.addParagraphTextItem()
  .setTitle('Desenvolvimento em soft-skills')
  .setValidation(null);

  // Adiciona pergunta de feedback nas hard-skills
  form.addParagraphTextItem()
  .setTitle('Desenvolvimento em hard-skills')
  .setValidation(null);
}
