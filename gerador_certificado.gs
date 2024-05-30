function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Gerador de certificados');
  menu.addItem('Gerar certificados', 'criarCertificados')
  menu.addItem('Enviar certificados por e-mail', 'enviarEmails')
  menu.addToUi();
  
}


function criarCertificados(){
  var ss = SpreadsheetApp.getActive()
  var sheetAssociados = ss.getSheetByName('Associados')
  // var valuesAssociados = sheetAssociados.getDataRange().getValues();
  var valuesAssociadosRaw = sheetAssociados.getRange(12, 2, sheetAssociados.getLastRow()-11, 7).getValues();
  var valuesAssociados = sheetAssociados.getRange(12, 2, sheetAssociados.getLastRow()-11, 7).getValues();

  // 1. Ordenar associados por nome
  valuesAssociados.sort(function(x,y){
        var xp = x[0];
        var yp = y[0];
        return xp == yp ? 0 : xp < yp ? -1 : 1;
      });

  
  const TIPO = ss.getRange("'Configurações e orientações'!C16").getValues()[0][0]
  const NOME_CURSO_TECNICO = ss.getRange("'Configurações e orientações'!C18").getValues()[0][0]
  const LINHA = ss.getRange("'Configurações e orientações'!C20").getValues()[0][0]
  const RAMO = ss.getRange("'Configurações e orientações'!C22").getValues()[0][0]
  const DIRETOR = ss.getRange("'Configurações e orientações'!C24").getValues()[0][0]
  const LOCAL = ss.getRange("'Configurações e orientações'!C27").getValues()[0][0]
  const DATA = ss.getRange("'Configurações e orientações'!C29").getValues()[0][0]
  const N_CURSO = ss.getRange("'Configurações e orientações'!C32").getValues()[0][0]
  const ID_CERTIFICADO = ss.getRange("'Configurações e orientações'!C35").getValues()[0][0]
  const ID_DESTINO = ss.getRange("'Configurações e orientações'!C37").getValues()[0][0]
  const CURR_DATE = new Date();

  var CURSO = null
  if (TIPO == "Curso Técnico"){
    CURSO = TIPO + ' - ' + NOME_CURSO_TECNICO
  }
  else if (LINHA == "Dirigente"){
    CURSO = LINHA
  }
  else{
    CURSO = LINHA + " - Ramo " + RAMO
  }



  var destinationFolder = DriveApp.getFolderById(ID_DESTINO)
  const googleSlideTemplate = DriveApp.getFileById(ID_CERTIFICADO);

  var reprovados = 0
  valuesAssociados.forEach(function(row, index){

    // 0. Verificar se o associado foi aprovado
    if(row[5] == 'Aprovado'){

      // 1. Criar cópia do documento
      if(TIPO == 'Curso Preliminar'){
        copyTemplate = googleSlideTemplate.makeCopy(`Certificado ${TIPO} -  ${row[0]}`, destinationFolder)
      }
      else if(TIPO == 'Curso Técnico'){
        copyTemplate = googleSlideTemplate.makeCopy(`Certificado ${TIPO} ${NOME_CURSO_TECNICO} -  ${row[0]}`, destinationFolder)
      }
      else if(LINHA == 'Dirigente'){
        copyTemplate = googleSlideTemplate.makeCopy(`Certificado ${TIPO} ${LINHA}  -  ${row[0]}`, destinationFolder)
      }
      else{
        copyTemplate = googleSlideTemplate.makeCopy(`Certificado ${TIPO} ${LINHA} Ramo ${RAMO} -  ${row[0]}`, destinationFolder)
      }

      const presentation  = SlidesApp.openById(copyTemplate.getId())

      // 2. preencher cópia do documento com dados do associado

      
      var temp_i = index - reprovados + 1
      var n_certificado = temp_i + "." + N_CURSO + "/" + (''+CURR_DATE.getFullYear()).substr(2)

      presentation.replaceAllText('<<associado>>', row[0]);
      presentation.replaceAllText('<<data>>', DATA);
      presentation.replaceAllText('<<n_certificado>>', n_certificado);
      presentation.replaceAllText('<<diretor>>', DIRETOR);
      presentation.replaceAllText('<<local>>', LOCAL);

      if(TIPO != "Curso Intermediário"){
        presentation.replaceAllText('<<curso>>', CURSO);
      }


      // 3. salvar número do certificado
      var index_escrita = null
      valuesAssociadosRaw.forEach(function(row_raw, index_raw){
        if (row_raw[0] == row[0]){
          index_escrita = index_raw
        }
      })

      sheetAssociados.getRange(12+index_escrita, 8).setValue(n_certificado)

    } 
    else{
      reprovados = reprovados + 1
      var index_escrita = null
      valuesAssociadosRaw.forEach(function(row_raw, index_raw){
        if (row_raw[0] == row[0]){
          index_escrita = index_raw
        }
      })

      sheetAssociados.getRange(12+index_escrita, 8).setValue('-')
    }
  })
}

function enviarEmails(){
  var ss = SpreadsheetApp.getActive()
  var sheetAssociados = ss.getSheetByName('Associados')
  var valuesAssociados = sheetAssociados.getRange(12, 2, sheetAssociados.getLastRow()-11, 7).getValues();


  
  const TIPO = ss.getRange("'Configurações e orientações'!C16").getValues()[0][0]
  const NOME_CURSO_TECNICO = ss.getRange("'Configurações e orientações'!C18").getValues()[0][0]
  const LINHA = ss.getRange("'Configurações e orientações'!C20").getValues()[0][0]
  const RAMO = ss.getRange("'Configurações e orientações'!C22").getValues()[0][0]
  const RESPONSAVEL_ENVIO = ss.getRange("'Configurações e orientações'!C39").getValues()[0][0]
  const FUNCAO_RESPONSAVEL = ss.getRange("'Configurações e orientações'!C41").getValues()[0][0]
  const DISTRITO = ss.getRange("'Configurações e orientações'!C43").getValues()[0][0]
  
  const ID_DESTINO = ss.getRange("'Configurações e orientações'!C37").getValues()[0][0]


  var folder = DriveApp.getFolderById(ID_DESTINO)

  var CURSO = null
  if (TIPO == 'Curso Preliminar'){
    CURSO = TIPO + ' - ' + NOME_CURSO_TECNICO
  }
  else if (TIPO == 'Curso Técnico'){
    CURSO = TIPO
  }
  else if (LINHA == "Dirigente"){
    CURSO = TIPO + " " + LINHA
  }
  else{
    CURSO = TIPO + " " + LINHA + " - Ramo " + RAMO
  }

  

  valuesAssociados.forEach(function(row, index){
    Logger.log(row)

    if (row[5] == "Aprovado"){
      Logger.log('Aprovado')

      var fileName = null
      if(TIPO == 'Curso Preliminar'){
        files = folder.getFilesByName(`Certificado ${TIPO} -  ${row[0]}`)
      }
      else if(TIPO == 'Curso Técnico'){
        files = folder.getFilesByName(`Certificado ${TIPO} ${NOME_CURSO_TECNICO} -  ${row[0]}`)
      }
      else if(LINHA == 'Dirigente'){
        files = folder.getFilesByName(`Certificado ${TIPO} ${LINHA}  -  ${row[0]}`)
      }
      else{
        Logger.log('Escotista')
        files = folder.getFilesByName(`Certificado ${TIPO} ${LINHA} Ramo ${RAMO} -  ${row[0]}`)
      }

      Logger.log(`Certificado ${TIPO} ${LINHA} Ramo ${RAMO} -  ${row[0]}`)

      while (files.hasNext()) {
        var file = files.next();
      }

      fillePDF = file.getAs('application/pdf')

      var templateAssociado = HtmlService.createTemplateFromFile('template_email');
      templateAssociado.associado = row[0];
      
      templateAssociado.responsavel = RESPONSAVEL_ENVIO;
      templateAssociado.funcao = FUNCAO_RESPONSAVEL;
      templateAssociado.distrito = DISTRITO;
      templateAssociado.curso = CURSO;
      
      var messageAssociado = templateAssociado.evaluate().getContent();

      if(row[4] != ""){
        MailApp.sendEmail({
          to: row[4],
          subject: `Seu certificado do ${CURSO} chegou!`,
          htmlBody: messageAssociado,
          attachments: [fillePDF]
        });

      }
    }
  })
}
