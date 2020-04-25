/* ********************************************  Inicio Nova Compra ******************************************* */
//Reformulação
function reformularCompraNova(){
  
  var spreadsheet = SpreadsheetApp.getActive();
  
  spreadsheet.getRange('D4').setBackground('#76a5af').clearDataValidations().setFormula('=IF(G5="";"";MAX(\'Compras Dados\'!A2:A)+1)');
  spreadsheet.getRange('D5').setFormula('=IF(G5="";"";IF(COUNTIF(\'Compras Dados\'!C2:C;G5) >= 1;LOOKUP(G5;\'Compras Dados\'!C2:C;\'Compras Dados\'!B2:B);MAX(\'Compras Dados\'!B2:B)+1))');
  
  spreadsheet.getRange('H6').setFormula('=IF(G5="";"";Today())');
  
  spreadsheet.getRange('K7').setFormula('=IF(K6="";"";L7)');
  spreadsheet.getRange('K8').setFormula('=IF(K7="";""; K6*K7)');
  

};


function modoNovaCompra() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  //Formulação
  spreadsheet.getRange('AL3').setValue(1);
  spreadsheet.getRange('D1').setValue("Novo");
  //Limpar
  spreadsheet.getRangeList(['G5','H6:H8','K4:K8','K25','G11:M15']).clear({contentsOnly: true, skipFilteredRows: true});
  
 // spreadsheet.getRange('AN3').setFormula('=IF(G5="";"";QUERY(\'Clientes Dados\'!A:M; "SELECT * WHERE \'"&G5&"\' = C "))');
  
  reformularCompraNova();
  
  spreadsheet.getRange('G18').setFormula('=IF(G11="";"";COUNTA(G11:G16))');
  spreadsheet.getRange('I18').setFormula('=IF(G11="";"";SUM(J11:J16))');
  spreadsheet.getRange('K18').setFormula('=IF(G11="";"";SUMPRODUCT(J11:J16;K11:K16))');
//spreadsheet.getRange('L18').setFormula('=IF(G11="";"";SUM(M11:M16))');
  
  spreadsheet.getRange('G25').setFormula('=IF(G5="";"";H6)');
  spreadsheet.getRange('I25').setFormula('=IF(K18="";"";K18)');
  spreadsheet.getRange('M25').setFormula('=IF(I25="";"";K18-I25)');
  spreadsheet.getRange('H22').setFormula('=IF(G18="";"";MAX(\'Compras Dados\'!M2:M)+1)'); //ID Pagamento

   spreadsheet.getRange('G5').activate();
  
};

/* ************************ Salvar Compra ******************** */

// Função auxiliar de salvarCompra
function salvarCompra2(x, idCompra){
  
  var spreadsheet = SpreadsheetApp.getActive();
  var DadosCompras = spreadsheet.getSheetByName('Compras Dados');
  var Form = spreadsheet.getSheetByName('Compras');
  
   var values = [[idCompra,                           // ID Compra
                   Form.getRange('D5').getValue(),    // ID Fornecedor
                   Form.getRange('G5').getValue(),    // Fornecedor
                   Form.getRange('H6').getValue(),    // Data Compra
                   Form.getRange('H7').getValue(),    // Motorista
                   Form.getRange('H8').getValue(),    // Placa
                   Form.getRange(x,7).getValue(),     // ID Produto
                   Form.getRange(x, 8).getValue(),    // Marca
                   Form.getRange(x, 9).getValue(),    // Produto
                   Form.getRange(x, 10).getValue(),   // Quantidade
                   Form.getRange(x, 11).getValue(),   // Custo de Compra
                   Form.getRange(x, 12).getValue(),   // Total de Compra
                   Form.getRange('H22').getValue(),   // ID Recebimento
                   Form.getRange('G25').getValue(),   // Data Recebimento
                   Form.getRange('K25').getValue(),   // Forma de Pagamento
                   Form.getRange('I25').getValue(),   // Valor recebido
                   Form.getRange('M25').getValue()]]; // Restante
  
  return DadosCompras.getRange(DadosCompras.getLastRow()+1,1,1,17).setValues(values);

};

// Função condutora para salvar Compra

function SalvarCompra() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName('Compras');
  var idCompra = Form.getRange('D4').getValue();
  var quantLinhasProd = spreadsheet.getRange('AJ3').getValue();
  
  if (spreadsheet.getRange('AK3').getValue() > 0 ) 
  {
    Browser.msgBox("Erro", "Necessário preencher todos os campos referente ao cliente, produto e recebimentos! ", Browser.Buttons.OK)
  }
  
  else{
    
       if(quantLinhasProd == 1){
         salvarCompra2(11, idCompra);
      
           }else{
 
           } if (Form.getRange('G12').getValue() != "") {
                 salvarCompra2(12, idCompra);
           } if (Form.getRange('G13').getValue() != "") {
                 salvarCompra2(13, idCompra);
           } if (Form.getRange('G14').getValue() != "") {
                 salvarCompra2(14, idCompra);
           } if (Form.getRange('G15').getValue() != "") {
                 salvarCompra2(15, idCompra);
           } if (Form.getRange('G16').getValue() != "") {
                 salvarCompra2(16, idCompra);
           }
    
                 limparProdCompras();
                 Browser.msgBox("Informativo", "Registro salvo com sucesso!", Browser.Buttons.OK)
                 reformularCompraNova();
                 spreadsheet.getRange('G5').activate();
    
       } 
         
};

/* ******************************************** Término Nova Compra ******************************************* */



/* ******************************************* Início Deletar Compras ********************************************** */

//Modo Deletar

function reformularDeletarCompra(){
  
  var spreadsheet = SpreadsheetApp.getActive();
 
  spreadsheet.getRange('H6').setFormula('=IF(D4="";"";BI5)'); //Data Compra
  spreadsheet.getRange('H7').setFormula('=IF(D4="";"";BJ5)'); //Entregador
  spreadsheet.getRange('H8').setFormula('=IF(D4="";"";BK5)'); //Veículo
  
  spreadsheet.getRange('K25').setFormula('=IF(D4="";"";BT5)'); //Forma de Pagamento
  
//spreadsheet.getRange('K7').setFormula('=IF(K6="";"";L7)');  //Preço de Compra
//spreadsheet.getRange('K8').setFormula('=IF(K7="";""; K6*K7)');// Total de Compra
  
  //Formular aŕea de produto
       
   var values = [['=IF($D$4="";"";BL5)','=IF($D$4="";"";BM5)','=IF($D$4="";"";BN5)','=IF($D$4="";"";BO5)','=IF($D$4="";"";BP5)','=IF($D$4="";"";BQ5)'],
                 ['=IF($D$4="";"";BL6)','=IF($D$4="";"";BM6)','=IF($D$4="";"";BN6)','=IF($D$4="";"";BO6)','=IF($D$4="";"";BP6)','=IF($D$4="";"";BQ6)'],
                 ['=IF($D$4="";"";BL7)','=IF($D$4="";"";BM7)','=IF($D$4="";"";BN7)','=IF($D$4="";"";BO7)','=IF($D$4="";"";BP7)','=IF($D$4="";"";BQ7)'],
                 ['=IF($D$4="";"";BL8)','=IF($D$4="";"";BM8)','=IF($D$4="";"";BM8)','=IF($D$4="";"";BO8)','=IF($D$4="";"";BP8)','=IF($D$4="";"";BQ8)'],
                 ['=IF($D$4="";"";BL9)','=IF($D$4="";"";BM9)','=IF($D$4="";"";BN9)','=IF($D$4="";"";BO9)','=IF($D$4="";"";BP9)','=IF($D$4="";"";BQ9)'],
                 ['=IF($D$4="";"";BL10)','=IF($D$4="";"";BM10)','=IF($D$4="";"";BN10)','=IF($D$4="";"";BO10)','=IF($D$4="";"";BP10)','=IF($D$4="";"";BQ10)']];

  spreadsheet.getRange('G11:L16').setValues(values);
  
}


function modoDeletarCompra(){

  var spreadsheet = SpreadsheetApp.getActive(); 
  spreadsheet.getRange('AL3').setValue(3);
  spreadsheet.getRange('D1').setValue("Deletar");
//Query consulta Compra
  spreadsheet.getRange('BF4').setFormula('=IF(G5="";QUERY(\'Compras Dados\'!A:R;"SELECT *");IF(D4="";QUERY(\'Compras Dados\'!A:R;"SELECT * WHERE "&D5&" = B");QUERY(\'Compras Dados\'!A:R;"SELECT * WHERE "&D5&" = B AND "&D4&" = A")))');
//Query ID Fornecedor
  spreadsheet.getRange('AN3').setFormula('=IF(G5="";"";QUERY(\'Compras Dados\'!A:C; "SELECT * WHERE \'"&G5&"\' = C "))');
//Limpar
  spreadsheet.getRangeList(['D4','G5','H6:H8','K4:K8','K25','G11:M15']).clear({contentsOnly: true, skipFilteredRows: true});
  
  spreadsheet.getRange('D4').setBackground('#ffffff').activate().setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false).requireValueInRange(spreadsheet.getRange('\'Compras\'!$BZ$5:$BZ'), true).build()); //ID Compra
  
  spreadsheet.getRange('D5').setFormula('=IF(G5="";"";AN4)'); //ID Cliente
//spreadsheet.getRange('D6').setFormula('=IF(D4="";"";BH5)'); //Canal de Compra
  
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('\'Compras Dados\'!$C$2:$C'), true).build());;

  reformularDeletarCompra();
  
  spreadsheet.getRange('H22').setFormula('=IF(D4="";"";BR5)'); //ID Pagamento
  
  spreadsheet.getRange('G25').setFormula('=IF(D4="";"";BS5)'); //Data Pagamento
  spreadsheet.getRange('I25').setFormula('=IF(D4="";"";BU5)'); //Valor Pago
  spreadsheet.getRange('K25').setFormula('=IF(D4="";"";BT5)'); //Forma de Pagamento
  spreadsheet.getRange('M25').setFormula('=IF(D4="";"";BV5)'); //Restante
  
  spreadsheet.getRange('G5').activate();

}

//**** Consolidar Deletar Compras

function DeletarCompra(){
  
  var spreadsheet = SpreadsheetApp.getActive();
  

  if (spreadsheet.getRange('AK3').getValue() > 0 ) {
    
    Browser.msgBox("Erro", "Necessário preencher os campos com ' * ' ", Browser.Buttons.OK);
    
  }else{

  var Compras = spreadsheet.getSheetByName('Compras');
//var Pesquisa = Compras.getRange('D4').getValue();
  var ComprasDados = spreadsheet.getSheetByName('Compras Dados');
//var LocalPesquisa = ComprasDados.getRange(2, 1, ComprasDados.getLastRow()).getValues();
//var Resultado = LocalPesquisa.findIndex(Pesquisa);
//var LINHA = Resultado + 2;
  var LINHA = spreadsheet.getRange('AI3').getValue();
  var quantLinhas = spreadsheet.getRange('AJ3').getValue();
  
  //ComprasDados.deleteRow(LINHA);
    
    ComprasDados.deleteRows(LINHA, quantLinhas);
    //Limpar
    spreadsheet.getRangeList(['D4','G5']).clear({contentsOnly: true, skipFilteredRows: true});   
    
    Browser.msgBox("Informativo", "Registro deletado!", Browser.Buttons.OK);
  
   //Reformular
   reformularDeletarCompra();

  spreadsheet.getRange('G5').activate();
}
};

/* *********************************************  Término Compras ********************************************** */



/* ******************************************** Função Inserir produto ***************************************** */
function inserirProdutoCompra2(x){
    
  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName('Compras');

    var values = [[Form.getRange('L5').getValue(),    // ID Produto
                   Form.getRange('K4').getValue(),    // Marca
                   Form.getRange('K5').getValue(),    // Produto
                   Form.getRange('K6').getValue(),    // Quantidade
                   Form.getRange('L7').getValue(),    // Custo de Compra
                 //Form.getRange('K7').getValue(),    // Preço de Compra                   
                   Form.getRange('K8').getValue()]];  // Total
    
  return Form.getRange(x).setValues(values), spreadsheet.getRangeList(['K4:K6']).clear({contentsOnly: true, skipFilteredRows: true}),
   spreadsheet.getRange('K25').activate();

};

function inserirProdutoCompra(){

var spreadsheet = SpreadsheetApp.getActive(); 
  
   
  if(spreadsheet.getRange('AM3').getValue() > 0){
   Browser.msgBox("Erro:","Necessário preencher os campos de produto!",Browser.Buttons);
    spreadsheet.getRange('K4').activate();

  }else{
  
  if (spreadsheet.getRange('G11').getValue() == ""){
      inserirProdutoCompra2('G11:L11');
    
  }else if (spreadsheet.getRange('G12').getValue() == ""){
            inserirProdutoCompra2('G12:L12');
    
  }else if (spreadsheet.getRange('G13').getValue() == ""){
            inserirProdutoCompra2('G13:L13');
    
    
  }else if (spreadsheet.getRange('G14').getValue() == ""){
            inserirProdutoCompra2('G14:L14');
    
    
  }else if (spreadsheet.getRange('G15').getValue() == ""){
            inserirProdutoCompra2('G15:L15');
    
    
  }else if (spreadsheet.getRange('G16').getValue() == ""){
            inserirProdutoCompra2('G16:L16');
    
           }else {
                  Browser.msgBox("Erro:","Todas as linhas foram preenchidas, finalize a Compra!",Browser.Buttons);
                  }
                    //spreadsheet.getRange('K8').setFormula('=IF(K7="";""; K6*K7)');
                    //spreadsheet.getRange('K7').setFormula('=IF(K6="";"";L7)');
          }
};

//******************    Finalizador   ******************************************************************


function FinalizadorCompra(){

  var spreadsheet = SpreadsheetApp.getActive();

  if(spreadsheet.getRange('AL3').getValue() == 1){
     SalvarCompra();
    
  }else if(spreadsheet.getRange('AL3').getValue() == 2){
           EditarCompra();
        
   }else{
         DeletarCompra();  
         } 

}

//*******************************************************************************************************

/*  ---------  Auxilio de Finalizadores ----------    */

// Declaração 
Array.prototype.findIndex = function(Procura){
    if (Procura == "") return false;
    for (var i = 0; i < this.length; i++)
      if (this[i] == Procura) return i;
      return -i;
};

/* ******************************************** Limpar  ******************************************* */
function limparProdCompras(){

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRangeList(['G5','G11:M16','H5:H8','K25']).clear({contentsOnly: true, skipFilteredRows: true}); 

};


/* ************************************************************************************************** */


function pagReltorioCompras(){

  var spreadsheet = SpreadsheetApp.getActive();
  //var compras = spreadsheet.getSheetByName('Compras');
  
 spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Relatório Compras'), true);
  
  

};

function relatorioCompras(){

  var spreadsheet = SpreadsheetApp.getActive();
  var url = "https://datastudio.google.com/embed/reporting/a4d2055e-8a6f-47b4-9f28-650ff423fa5b/page/feyJB"
  var html = "<script> window.open('"+ url + "');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html);
  
  SpreadsheetApp.getUi().showModalDialog(userInterface, "Relatório de Compras...");
}
/* Exemplo de criar abrir relotario por uma caixa de dialogo ou direto

function relatoriosComprasDialog() {
  
  var url = 'https://datastudio.google.com/embed/reporting/a4d2055e-8a6f-47b4-9f28-650ff423fa5b/page/feyJB';
  var name = 'Compras';
  var url2 = 'https://datastudio.google.com/embed/reporting/a4d2055e-8a6f-47b4-9f28-650ff423fa5b/page/feyJB';
  var name2 = 'Compras';  
  var html = '<html><body><a href="'+url+'" target="blank" onclick="google.script.host.close()">'+name+'</a> <br><br/><a href="'+url2+'" target="blank" onclick="google.script.host.close()">'+name2+'</a></body></html>';
  var ui = HtmlService.createHtmlOutput(html)
  SpreadsheetApp.getUi().showModelessDialog(ui,"Relatórios de Compras");
}

function relatorioCompras(){

  var spreadsheet = SpreadsheetApp.getActive();
  var url = "https://datastudio.google.com/embed/reporting/a4d2055e-8a6f-47b4-9f28-650ff423fa5b/page/feyJB"
  var html = "<script> window.open('"+ url + "');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html);
  
  SpreadsheetApp.getUi().showModalDialog(userInterface, "Relatório de Compras...");
}

*/