var geralRegex = new RegExp("De 1 a 10, o quanto você considera que a Inderios está atendendo suas expectativas quanto a (.*)")
var macroareasRegex = new RegExp("De 1 a 10, como você avalia seu conhecimento em (.*)")
var skillsRegex = new RegExp('De 1 a 10, como você se avalia em (.*)')
var preMedraRegex = new RegExp('Autoavaliação pré-medra: (.*)')

/**
 * Retorna uma lista de vetores do tipo: 
 * ['1','2','Hard Skill','Excel',2,4]
 * em que
 * ['ID pessoa','Nº resposta','Tipo','Expectativa','Nota pré-medra','Nota resposta']
 * 
 * tipo pode assumir "Geral","Macroareas","Soft Skills" ou "Hard Skills"
 */
function getRespostas(id,idPessoa){

  // DEFINIÇÃO DAS VARIÁVEIS
  var form = FormApp.openById(id);
  var formResponses = form.getResponses();
  var itens = form.getItems();
  var respostas = [];
  var feedbacks = [];

  // Percorre todas as respostas do formulario
  for (var resposta = 0; resposta < formResponses.length; resposta++) {
    // Armazena todos os itens de uma resposta do forms
    var formResponse = formResponses[resposta];

    // Cria estrutura na qual serao armazenados os dados
    var dados = {'tipo':'Geral'};

    // Percorre todos os itens (questoes) para a resposta do formulario
    // (cabecalhos, perguntas)
    for (var item = 0; item < itens.length; item++) {

      // Se o item é um cabecalho de secao, altera a variavel "tipo" 
      if(itens[item].getType() == "SECTION_HEADER"){dados['tipo'] = itens[item].getTitle();}

      // Se o item for a pergunta de confirmacao do nome, nao faz nada
      else if(itens[item].getIndex() == 0){}


      // Se estiver na secao de feedbaco
      else if(dados.tipo == "Feedback"){
        try{
        feedbacks.push(formResponse.getResponseForItem(itens[item]).getResponse());
        }catch(err){}
      }

      // Se nao, armazena os dados do item
      else{
        try{
          var pergunta = itens[item];

          // Armazena qual eh o numero da resposta do forms
          dados['numRealizacao'] = (resposta + 1);

          // Armazena qual eh o texto da pergunta
          texto = pergunta.getTitle();

          if(dados.tipo == 'Geral'){
            dados['desejo'] = geralRegex.exec(texto)[1];
            dados['preMedra'] = null;
          }else{
            descricao = pergunta.getHelpText();
            dados['preMedra'] = preMedraRegex.exec(descricao)[1];

            if(dados.tipo == 'Macroareas'){dados['desejo'] = macroareasRegex.exec(texto)[1];}

            else{dados['desejo'] = skillsRegex.exec(texto)[1];}
          }

          // Armazena a nota dada no forms
          dados['nota'] = formResponse.getResponseForItem(itens[item]).getResponse();

          respostas.push([idPessoa, dados.numRealizacao, dados.tipo, dados.desejo, dados.preMedra, dados.nota]);
        }
        catch(err){
          Logger.log(err);
          Logger.log(texto)
        }
      }
    }
  }
  //return respostas;
  return {
    'respostas':respostas,
    'feedbacks': feedbacks
  }
}

/**
 * Registra todas as respostas do formulario em uma planilha
 */
function getAll(){
  var info = getInfo();
  var folderID = getFolderId(info.pasta_destino);
  var files = DriveApp.getFolderById(folderID).getFiles();
  var pessoa = 0;
  var dados = [];

  // Percorre todos os arquivos da pasta
  while (files.hasNext()) {
    pessoa += 1;
    var file = files.next();
    Logger.log(file.getName());

    // Armazena as respostas do forms
    dados.push(getRespostas(file.getId(),pessoa));
  }
  console.log(dados[0].respostas)

  // Registra os dados na planilha de destino
  registraDados(dados)
}

function getFolderId(url){
  return url.split('/').slice(-1);
}

function registraDados(dados){
  var info = getInfo();
  var spreadsheet = SpreadsheetApp.openByUrl(info.link_respostas);
  var header1 = [['ID pessoa','Nº resposta','Tipo','Expectativa','Nota pré-medra','Nota resposta']];
  var header2 = [['Expectativa ao entrar', 'Macroareas', 'Soft-skills', 'Hard-skills']];

  // Cria a aba de destino dos dados
  spreadsheet.insertSheet().setName('Dados realização');
  var sheet = spreadsheet.getSheetByName('Dados realização');
  
  //Preeche cabecalho
  sheet.getRange(1,1,1,header1[0].length).setValues(header1);

  console.log('DADOS:\n' +dados.length)
  //console.log('dados.respostas:\n' + dados.respostas)

  // Percorre todas as respostas do formulario registrando linha por linha
  for(var formulario = 0; formulario < dados.length; formulario++){
    //for(var response = 0; )
    var linha = sheet.getLastRow() + 1;
    sheet.getRange(linha,1,dados.respostas[formulario].length,header1[0].length).setValues(dados.respostas[formulario]);
  }

  // Cria a aba de destino de feedback
  spreadsheet.insertSheet().setName('Feedbacks');
  var sheet = spreadsheet.getSheetByName('Feedbacks');

  //Preeche cabecalho
  sheet.getRange(1,1,1,header2[0].length).setValues(header2);

  // Registra todos os feedbacks
  for(var f = 0; f < dados.feedbacks.length; f++){
    var linha = sheet.getLastRow() + 1;
    
    for(col = 0; col <= 3; col++){
      sheet.getRange(linha, col).setValue(dados.feedbacks[f]);
    }
  }
}
