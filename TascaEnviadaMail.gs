function onFormSubmit(e) 
{
  var emails = GmailApp.getInboxThreads();
  //recupero fecha de la última ejecución
  var fecha = new Date(PropertiesService.getScriptProperties().getProperty('fecha'));
  var sheet = fullaActiva2(); //selecciono document excel

  for(var em = 0 ; em < emails.length && em<20 ; ++em )
  {
     var email = emails[em];
     //const date1 = new Date('October 27, 2023 00:00:00 -0500');

     if ( email.isInInbox()  && fecha.valueOf()<email.getMessages()[0].getDate().valueOf() )
     { //si el email es posterior a la última revisió?
        date2 = email.getMessages()[0].getDate();

         
        var userEmail = email.getMessages()[0].getFrom();

        //recupero el assumpte
        var asunto = email.getMessages()[0].getSubject();
        
        //funció que retorna el email sense basura entremig
        var adrecaEmail = trobarEmail(userEmail); // L'adreça de correu electrònic es troba a l'índex 1 de la coincidència

        //inicialitzo variables, per saber si trobo el mail i el assumpte
        var fila = -1, columna=-1;

        fila = trobarFila(sheet,adrecaEmail);

        columna = -1;
        if (fila!=-1)
        {
          columna = trobarCol(sheet,asunto);
        }
        if((fila!== -1) && (columna!==-1))
        { //trobada entrega i mail, la marco com assignada.
          sheet.getRange(fila,columna).setValue("X");
        }
        if (em==0) //sol el primer cop cambio la data del sistema
        { //pel proper cop que s'executi que sol agafi mails, posteriors a l'execució
          PropertiesService.getScriptProperties().setProperty('fecha', new Date());
        }
     }
  }
}

/**
 *  Busca a partir del mail rebut, la fila on es troba aquest alumne.
 */
function trobarFila(sheet, mail)
{

  for (var fila = 1; fila <= sheet.getLastRow(); fila++) {
    for (var columna = 4; columna <= 5; columna++) { //limito les columnes perque sol vull que miri unes específiques.
        Logger.log(sheet.getRange(fila,columna).getValues() + "====" + mail);
      if (sheet.getRange(fila,columna).getValues().valueOf() == mail.valueOf()) {
        return fila;
      }
    }
  }

  return -1;
}

/**
 * Busca dins les files possibles, i dins tota la columna
 * el titol del assumpte.
 */
function trobarCol(sheet, assumpte)
{ //busco sol dins les columnes triades.
  //Logger.log("trobarCela " + sheet.getLastRow());
  for (var fila = 3; fila <= 5; fila++) { //limito les files perque sol vull que miri unes específiques.
    for (var columna = 1; columna <= sheet.getLastColumn(); columna++) {
      var info = sheet.getRange(fila,columna).getValues();   
      info = info.toString().toLowerCase();
      //if (sheet.getRange(fila,columna).getValues() == assumpte.valueOf()) {
      if(info.indexOf(assumpte.toString().toLowerCase())>=-1) //existe
      {
        return columna;
      }
    }
  }
  return -1;
}

/**
 * Selecciones el document i la fulla de google Calc.
 */
function fullaActiva2()
{
    var spreadsheet = SpreadsheetApp.openById('direcció google sheets');
    var sheet = spreadsheet.getSheetByName("nom fulla");
    return sheet;

}

/**
 * Trobes email, dins tota la informació del usuari que retorna el mètode userFrom
 */
function trobarEmail(userFrom)
{
        var regex = /\<(.*?)\>/; // Expressió regular per trobar l'adreça de correu electrònic entre els símbols <>
        var coincidencia = regex.exec(userFrom); // Troba la coincidència en la cadena del remitent
        
        if (coincidencia && coincidencia.length > 1) {
          var adrecaEmail = coincidencia[1]; // L'adreça de correu electrònic es troba a l'índex 1 de la coincidència
          return adrecaEmail; 
        } else {
          return "";
        }
}
