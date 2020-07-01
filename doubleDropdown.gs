/*Apps Script para doble validación - adapted by dvfdez

1.Crear dos hojas

2.Establecer el nombre de la hoja principal de trabajo y colocarlo en la variable: mainDataSheetName

3.Establecer el nombre de la hoja que llevará la tabla con los datos de validación y colocarlo en la 
variable: validationDataSheetName

b.Crear la tabla principal de validación, colocando como encabezado los campos de validación principal

c.En cada columna anotar los datos aplicables que debe desplegar el validación secundaria 
ejemplo
Letras	Numeros	Simbolos
A       0       !
B       1       -
C       2       *
D       3       $
E       4       %

4.En la hoja principal establecer la validación principal de datos desde la fila 2 de la columna deseada,
seleccionando como rango de validación la fila de encabezados de la hoja que contiene la tabla de validación

5.Configurar variables principales en el script
*/

function onEdit(){
  //config
  var mainDataSheetName = "Datos";            //Nombre de la hoja con los datos principales
  var validationDataSheetName = "Validacion"; //Nombre de la hoja con tabla de validacion
  var firstValidationColumn = 3;              //Número de columna de la validación principal
  var secondValidationColumnOffset = 3;       //Número de columna para la validación secundaria
  var allowInvalid = false;                   //¿Permitir que se escriban datos fuera del rango de validación? true/false
  //end config
  
  //obtiene la hoja con datos de validacion
  var validationDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(validationDataSheetName);
  
  //obtiene la hoja activa
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  //obtiene la celda activa
  var activeCell = activeSheet.getActiveCell();
    
  //comprueba si se edita en la hoja de datos la columna de validacion principal a partir de la segunda fila
  if(activeSheet.getSheetName() === mainDataSheetName && activeCell.getColumn() === firstValidationColumn && activeCell.getRow() > 1){
    
    //elimina el contenido y validaciones presentes en la celda de la validación secundaria
    activeCell.offset(0, secondValidationColumnOffset).clearContent().clearDataValidations();
    
    //obtiene los valores actuales de la tabla de validación
    var validationDataTable = validationDataSheet.getRange(1, 1, 1, validationDataSheet.getLastColumn()).getValues();
    
    //determina el índice de columna a utilizar según lo ingresado en validación principal
    var index = validationDataTable[0].indexOf(activeCell.getValue()) + 1;
    
    //comprueba si el índice determinado es válido
    if(index != 0){
      
      //obtiene los datos del rango de validación secundaria según el índice
      var validationRange = validationDataSheet.getRange(2, index, validationDataSheet.getLastRow());
      
      //construye la regla de validación secundaria
      var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).setAllowInvalid(allowInvalid).build();
      
      //aplica la regla de validación en la fila correspondiente de la columna de validación secundaria
      activeCell.offset(0, secondValidationColumnOffset).setDataValidation(validationRule);
  
     }  
      
  }
  
}
