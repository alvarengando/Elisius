function modoSalvarRecebimento() {
   var spreadsheet = SpreadsheetApp.getActive(); 
  
  spreadsheet.getRange('AL3').setValue(1);
  spreadsheet.getRange('D1').setValue("Novo");
  spreadsheet.getRangeList(['F4:F9','C8']).clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C8').setBackground('#ffffff').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('!$C$18:$C'), true).build());;
  
   spreadsheet.getRange('C3').setFormula('=IF(F4="";"";MATCH(F4;\'Dados Clientes\'!B2:B;))');
   spreadsheet.getRange('C5').setBackground('#addac2').clearDataValidations().setFormula('=IF(F4="";"";MAX(\'Recebimentos Dados\'!A:A)+1)');
   spreadsheet.getRange('F4').activate().setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('\'Consignados Dados\'!$C$2:$C'), true).build()); 
   spreadsheet.getRange('F6').setFormula('=IF(C8="";""; VLOOKUP(C8;C18:D57;2))');
   spreadsheet.getRange('F7').setFormula('=IF(C8="";"";MAX(O5:O100)+1)');
   spreadsheet.getRange('F8').setFormula('=IF(C8="";""; VLOOKUP(C8;C18:F57;4))');
  
   spreadsheet.getRange('C17').setFormula( '=IF(F4="";""; QUERY(\'Consignados Dados\'!A:J;"SELECT A,F,G,H,I WHERE \'"&F4&"\' = C "))');
   spreadsheet.getRange('I4').setFormula('=IF(F4="";QUERY(\'Recebimentos Dados\'!A:I;"ORDER BY F DESC");IF(C8="";QUERY(\'Recebimentos Dados\'!A:I;"SELECT * WHERE \'"&F4&"\' = C ORDER BY F DESC"); QUERY(\'Recebimentos Dados\'!A:I;"SELECT * WHERE \'"&F4&"\' = C AND "&C8&" = E ORDER BY F DESC")))');

}




// Modo Editar 
function modoEditarRecebimento() {
   var spreadsheet = SpreadsheetApp.getActive(); 
  
  spreadsheet.getRange('AL3').setValue(2);
  spreadsheet.getRange('D1').setValue("Editar");
  spreadsheet.getRange('C8').setBackground('#addac2');
  spreadsheet.getRangeList(['F4:F9','C5']).clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C5').setBackground('#ffffff').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('!$I$5:$I100'), true).build());;
  
  spreadsheet.getRange('C8').clearDataValidations().setFormula('=IF(C5="";""; LOOKUP(C5;I5:M100;M5:M100))');
  
   spreadsheet.getRange('F4').activate().setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('\'Recebimentos Dados\'!$C$2:$C'), true).build());
  spreadsheet.getRange('F5').setFormula('=IF(C5="";""; LOOKUP(C5;I5:M100;L5:L100))');
  spreadsheet.getRange('F6').setFormula('=IF(C5="";""; LOOKUP(C5;I5:N100;N5:N100))');
  spreadsheet.getRange('F7').setFormula('=IF(C5="";""; LOOKUP(C5;I5:O100;O5:O100))');
  spreadsheet.getRange('F8').setFormula('=IF(C5="";""; LOOKUP(C5;I5:P100;P5:P100))');
  spreadsheet.getRange('F9').setFormula('=IF(C5="";""; LOOKUP(C5;I5:Q100;Q5:Q100))');
  
   spreadsheet.getRange('C17').setFormula( '=IF(F4="";""; QUERY(\'Consignados Dados\'!A:J;"SELECT A,F,G,H,I WHERE \'"&F4&"\' = C "))');
   spreadsheet.getRange('I4').setFormula('=IF(F4="";QUERY(\'Recebimentos Dados\'!A:I);QUERY(\'Recebimentos Dados\'!A:I;"SELECT * WHERE \'"&F4&"\' = C"))');

}

// Modo Deletar 
function modoDeletarRecebimento() {
   var spreadsheet = SpreadsheetApp.getActive(); 
  
  spreadsheet.getRange('AL3').setValue(3);
  spreadsheet.getRange('D1').setValue("Deletar");
  spreadsheet.getRange('C8').setBackground('#addac2');
  spreadsheet.getRangeList(['F4:F9','C5']).clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C5').setBackground('#ffffff').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('!$I$5:$I100'), true).build());;
  
  spreadsheet.getRange('C8').clearDataValidations().setFormula('=IF(C5="";""; LOOKUP(C5;I5:M100;M5:M100))');
  spreadsheet.getRange('F4').activate().setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('\'Recebimentos Dados\'!$C$2:$C'), true).build());
  spreadsheet.getRange('F5').setFormula('=IF(C5="";""; LOOKUP(C5;I5:M100;L5:L100))');
  spreadsheet.getRange('F6').setFormula('=IF(C5="";""; LOOKUP(C5;I5:N100;N5:N100))');
  spreadsheet.getRange('F7').setFormula('=IF(C5="";""; LOOKUP(C5;I5:O100;O5:O100))');
  spreadsheet.getRange('F8').setFormula('=IF(C5="";""; LOOKUP(C5;I5:P100;P5:P100))');
  spreadsheet.getRange('F9').setFormula('=IF(C5="";""; LOOKUP(C5;I5:Q100;Q5:Q100))');
 
   spreadsheet.getRange('C17').setFormula( '=IF(F4="";""; QUERY(\'Consignados Dados\'!A:J;"SELECT A,F,G,H,I WHERE \'"&F4&"\' = C "))');
   spreadsheet.getRange('I4').setFormula('=IF(F4="";QUERY(\'Recebimentos Dados\'!A:I);QUERY(\'Recebimentos Dados\'!A:I;"SELECT * WHERE \'"&F4&"\' = C"))');

}



//******************    Finalizador   ******************************************************************


function FinalizadorRecebimentos(){

  var spreadsheet = SpreadsheetApp.getActive();

  if(spreadsheet.getRange('AL3').getValue() == 1){
  
    SalvarRecebimento();
    
  }else if(spreadsheet.getRange('AL3').getValue() == 2){
  
  EditarRecebimentos();
  
  
  }else{
    
    DeletarRecebimentos();  
  
  } 


}


//*******************************************************************************************************
//*******************************************************************************************************



/*  ---------  Auxilio de Finalizadores ----------    */


// Declaração 
Array.prototype.findIndex = function(Procura){
    if (Procura == "") return false;
    for (var i = 0; i < this.length; i++)
      if (this[i] == Procura) return i;
      return -i;

  
  
};


//**** SALVAR Recebimentos


function SalvarRecebimento() {
  var spreadsheet = SpreadsheetApp.getActive();

  
  if (spreadsheet.getRange('AJ3').getValue() > 0 ) 
  {
    Browser.msgBox("Erro", "Necessário preencher os campos! ", Browser.Buttons.OK)
    
  }else{
  
  var RecebimentosDados = spreadsheet.getSheetByName('Recebimentos Dados');
  var Form = spreadsheet.getSheetByName('Recebimentos');

    var values = [[Form.getRange('C5').getValue(),
                   Form.getRange('C3').getValue(), //ID CLiente
                   Form.getRange('F4').getValue(),//Cliente
                   Form.getRange('F5').getValue(), //Data
                   Form.getRange('C8').getValue(), //ID Consignado
                   Form.getRange('F6').getValue(),  //Consignado
                   Form.getRange('F7').getValue(),
                   Form.getRange('F8').getValue(),
                   Form.getRange('F9').getValue()]];
    
    
  
  RecebimentosDados.getRange(RecebimentosDados.getLastRow()+1,1,1,9).setValues(values);
 
  spreadsheet.getRangeList(['F4:F9','C8']).clear({contentsOnly: true, skipFilteredRows: true});
 
   spreadsheet.getRange('C3').setFormula('=IF(F4="";"";MATCH(F4;\'Dados Clientes\'!B2:B;))');
   spreadsheet.getRange('C5').setBackground('#addac2').clearDataValidations().setFormula('=IF(F4="";"";MAX(\'Recebimentos Dados\'!A:A)+1)');
   spreadsheet.getRange('F6').setFormula('=IF(C8="";""; VLOOKUP(C8;C18:D57;2))');
   spreadsheet.getRange('F7').setFormula('=IF(C8="";"";MAX(O5:O100)+1)');
   spreadsheet.getRange('F8').setFormula('=IF(C8="";""; VLOOKUP(C8;C18:F57;4))');
   spreadsheet.getRange('F4').activate();
    
  Browser.msgBox("Informativo", "Registro salvo com sucesso!", Browser.Buttons.OK)

  }
  
 };



/// ***** Editar Recebimentos ******


function EditarRecebimentos(){
  
  
  var spreadsheet = SpreadsheetApp.getActive();

  
  if (spreadsheet.getRange('AJ3').getValue() == 0 ) {
    
  var Recebimentos = spreadsheet.getSheetByName("Recebimentos");
  var RecebimentosDados = spreadsheet.getSheetByName('Recebimentos Dados');
 
  var Pesquisa = Recebimentos.getRange('C5').getValue();
  
  Recebimentos.getActiveCell();
  
  var LocalPesquisa = RecebimentosDados.getRange(1, 1, RecebimentosDados.getLastRow()).getValues();
  var Resultado = LocalPesquisa.findIndex(Pesquisa);
  
  var LINHA = Resultado + 1;
 
 
  if (Resultado != -1) {
  
  
    RecebimentosDados.getActiveCell(); //Novamente?
    
   // ConsignadosDados.getRange(LINHA,ConsignadosDados.getLastRow()).setValues(Consignados.getRange(['D6','D10','I5:I12']).getValues());
    RecebimentosDados.getRange(LINHA,4).setValue(Recebimentos.getRange('F5').getValue());
    RecebimentosDados.getRange(LINHA,7).setValue(Recebimentos.getRange('F7').getValue());
    RecebimentosDados.getRange(LINHA,8).setValue(Recebimentos.getRange('F8').getValue());
    RecebimentosDados.getRange(LINHA,9).setValue(Recebimentos.getRange('F9').getValue());
    
    
    
    //Limpar
    spreadsheet.getRangeList(['F4:F9','C8','C5']).clear({contentsOnly: true, skipFilteredRows: true});    
    
    Browser.msgBox("Informativo", "Registro Alterado!", Browser.Buttons.OK);
   // Reformular
  spreadsheet.getRange('C8').clearDataValidations().setFormula('=IF(C5="";""; LOOKUP(C5;I5:M100;M5:M100))');
  spreadsheet.getRange('F5').setFormula('=IF(C5="";""; LOOKUP(C5;I5:M100;L5:L100))');
  spreadsheet.getRange('F6').setFormula('=IF(C5="";""; LOOKUP(C5;I5:N100;N5:N100))');
  spreadsheet.getRange('F7').setFormula('=IF(C5="";""; LOOKUP(C5;I5:O100;O5:O100))');
  spreadsheet.getRange('F8').setFormula('=IF(C5="";""; LOOKUP(C5;I5:P100;P5:P100))');
  spreadsheet.getRange('F9').setFormula('=IF(C5="";""; LOOKUP(C5;I5:Q100;Q5:Q100))');
   
   

  } else {

    Browser.msgBox("Pedido não Localizado!")

  }

  }else{
  
  
   Browser.msgBox("Erro", "Necessário preencher os campos!  ", Browser.Buttons.OK);
  
  }

}


//**** Deletar Recebimentos

function DeletarRecebimentos(){
  
  
  var spreadsheet = SpreadsheetApp.getActive();

  if (spreadsheet.getRange('AJ3').getValue() > 0 ) {
    
    Browser.msgBox("Erro", "Necessário preencher os campos! ", Browser.Buttons.OK);
    
  }else{
  

  var Recebimentos = spreadsheet.getSheetByName('Recebimentos');
  var Pesquisa = Recebimentos.getRange('C5').getValue();
  var RecebimentosDados = spreadsheet.getSheetByName('Recebimentos Dados');
  RecebimentosDados.getActiveCell();
  
  var LocalPesquisa = RecebimentosDados.getRange(2, 1, RecebimentosDados.getLastRow()).getValues();
  var Resultado = LocalPesquisa.findIndex(Pesquisa);
  
  var LINHA = Resultado + 2;
  
  
  if (Resultado != -1) {
  
  RecebimentosDados.deleteRow(LINHA);
    
  spreadsheet.getRangeList(['F4:F9','C5']).clear({contentsOnly: true, skipFilteredRows: true});

  spreadsheet.getRange('C8').clearDataValidations().setFormula('=IF(C5="";""; LOOKUP(C5;I5:M100;M5:M100))');
  spreadsheet.getRange('F5').setFormula('=IF(C5="";""; LOOKUP(C5;I5:M100;L5:L100))');
  spreadsheet.getRange('F6').setFormula('=IF(C5="";""; LOOKUP(C5;I5:N100;N5:N100))');
  spreadsheet.getRange('F7').setFormula('=IF(C5="";""; LOOKUP(C5;I5:O100;O5:O100))');
  spreadsheet.getRange('F8').setFormula('=IF(C5="";""; LOOKUP(C5;I5:P100;P5:P100))');
  spreadsheet.getRange('F9').setFormula('=IF(C5="";""; LOOKUP(C5;I5:Q100;Q5:Q100))');   
    Browser.msgBox("Informativo", "Registro deletado!", Browser.Buttons.OK);
 

  }
}
}



