const PLANILHA = 'https://docs.google.com/spreadsheets/d/1MvNgfmsKN7ZvMYP6Pr3SFJmYTFgEzL7AUuzFGitghHA/edit#gid=0';

/**
 * Retorna as informações necessárias para executar o script
 */
function getInfo(){
  var worksheet = SpreadsheetApp.openByUrl(PLANILHA).getSheetByName('Expectativas');
  var info = {};

  info['link_respostas'] = worksheet.getRange('C3').getValue();
  info['aba_respostas'] = worksheet.getRange('C4').getValue();
  info['pasta_destino'] = worksheet.getRange('C5').getValue();
  info['link_checks'] = worksheet.getRange('C6').getValue();
  info['num_pessoas'] = worksheet.getRange('C7').getValue();

  return info;
}

/**
 * Cria todos os formularios personalizados e coloca os links na planilha de verificacao
 */
function criaForms() {
  var info = getInfo();
  var worksheet = SpreadsheetApp.openByUrl(info.link_respostas);
  var sheet = worksheet.getSheetByName(info.aba_respostas);

  var nomes = getNames(2, 2, info.num_pessoas, sheet);
  var expectativas = getSkills(2, 3, info.num_pessoas, sheet);
  var softSkills = getSkills(2, 5, info.num_pessoas, sheet, sep=';');
  var hardSkills = getSkills(2, 19, info.num_pessoas, sheet, sep=';');
  var macroareas = ['Finanças', 'Mercado', 'Estratégia Organizacional', 'Marketing e vendas'];

  var notasMacroareas = sheet.getSheetValues(2,15,info.num_pessoas,4);
  var notasSoftSkills = sheet.getSheetValues(2,6,info.num_pessoas,5);
  var notasHardSkills = sheet.getSheetValues(2,20,info.num_pessoas,5);
  
  tabelaHeader(info.link_checks);

  for(var person=0; person < nomes.length; person++){
    var form = FormApp.create(nomes[person]);   //Cria formulário com o nome da pessoa

    moveFiles(form.getId(),info.pasta_destino);    //Move formulário para a pasta destino

    form.addMultipleChoiceItem()
      .setTitle('Você é ' + nomes[person] + '?')
      .setChoiceValues(['Sim']);    //Cria pergunta para confirmar se o forms respondido é o correto

    //Cria as perguntas de expectativa
    for(var expectativa = 0; expectativa < expectativas[person].length; expectativa++){
      form.addScaleItem()
        .setBounds(1,10)
        .setRequired()
        .setTitle('De 1 a 10, o quanto você considera que a Inderios está atendendo suas expectativas quanto a ' + expectativas[person][expectativa]);
    }
    
    //Adiciona seção e cria perguntas das macroareas
    form.addSectionHeaderItem().setTitle('Macroareas')
    for(var area = 0; area < macroareas.length; area++){
      form.addScaleItem()
        .setBounds(1,10)
        .setRequired()
        .setTitle('De 1 a 10, como você avalia seu conhecimento em ' + macroareas[area])
        .setHelpText('Autoavaliação pré-medra: ' + notasMacroareas[person][area]);
    }
    
    form.addSectionHeaderItem().setTitle('Soft Skills')
    for(var skill = 0; skill < softSkills[person].length; skill++){
      form.addScaleItem()
        .setBounds(1,10)
        .setRequired()
        .setTitle('De 1 a 10, como você se avalia em ' + softSkills[person][skill])
        .setHelpText('Autoavaliação pré-medra: ' + notasSoftSkills[person][skill]);
        //console.log(softSkills[person][skill]);
    }

    form.addSectionHeaderItem().setTitle('Hard Skills')
    for(var skill = 0; skill < hardSkills[person].length; skill++){
      form.addScaleItem()
        .setBounds(1,10)
        .setRequired()
        .setTitle('De 1 a 10, como você se avalia em ' + hardSkills[person][skill])
        .setHelpText('Autoavaliação pré-medra: ' + notasHardSkills[person][skill]);
        //console.log(hardSkills[person][skill])
    }

    colocarLink(info.link_checks, form.getPublishedUrl(), nomes[person]);
    console.log(nomes[person] + ' ok');
  }

  formataTabela(info.link_checks);
}

/**
 * Coloca os links e checkbox na planilha de verificacao
 */
function colocarLink(link_planilha_destino, link_forms, nome){
  var destino = SpreadsheetApp.openByUrl(link_planilha_destino);
  var currentLine = destino.getLastRow() + 1;

  destino.getRange('A'+currentLine).setValue(nome);
  destino.getRange('B'+currentLine).setValue(link_forms);
  destino.getRange('C'+currentLine).insertCheckboxes();
}

/**
 * Cria o cabecalho da planilha de verificacao
 */
function tabelaHeader(link_planilha_destino){
  var sheet = SpreadsheetApp.openByUrl(link_planilha_destino);
  sheet.getRange('A1:C1').setValues([['NOME', 'LINK', 'RESPONDIDO?']]);
}

/**
 * Formata a tabela da planilha de verificacao
 */
function formataTabela(link_planilha_destino){
  var sheet = SpreadsheetApp.openByUrl(link_planilha_destino);

  sheet.getDataRange().applyRowBanding();
  var banding = sheet.getDataRange().getBandings()[0];
  banding
    .setHeaderRowColor('#543396')
    .setFirstRowColor('#ffffff')
    .setSecondRowColor('#ebe2e9')
    .setFooterRowColor(null);
  sheet.getRange('A1:C1')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
}

function getNames(firstNameRow=2,firstNameColumn, numRows, sheet){
  var _names = sheet.getSheetValues(firstNameRow,firstNameColumn, numRows, 1);
  var names = [];

  for(var i = 0; i < numRows; i++){
    names[i] = _names[i][0].trim();
  }
  return names
}

function getSkills(firstRow=2, column, numRows,  sheet, sep = ','){
  var _skills = sheet.getSheetValues(firstRow, column, numRows, 1);
  var skills = [];

  for(var i = 0; i < numRows; i++){
    skills[i] = _skills[i][0].split(sep);
    for(var j in skills[i]){
      skills[i][j] = skills[i][j].trim()
    }
  }
  return skills
}

function getAutoavaliacao(skills, firstSkillColumn){
  var info = getInfo();
  var worksheet = SpreadsheetApp.openByUrl(info.link_respostas);
  var sheet = worksheet.getSheetByName(info.aba_respostas);
  var _notas = sheet.getSheetValues(2,6,22,5);

  console.log(sheet.getSheetValues(2,6,22,5))

  for(var i = 0; i < skills.length; i++){
     for(var j = 0; j < skills[i].length; j++){
       notas[i] = sheet.getRange(2 + i, firstSkillColumn + j);
     }
     console.log(notas[i]);
  }

  //console.log(sheet.getSheetValues('F7:J8'))
  console.log('Resultado: ' + notas);
}


function getNotas(row=2, columm, numRows, numColumns, sheet){
  var notas = sheet.getSheetValues(row,columm,numRows,numColumns);
  return notas;
}

function moveFiles(sourceFileId, targetFolderId) {
  var file = DriveApp.getFileById(sourceFileId);
  var folder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(folder);
}
