function inserirProdutoVenda2(x){

  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName('Vendas');
 
  var values = [[Form.getRange('L5').getValue(),    //ID Produto
                   Form.getRange('K4').getValue(),    // Marca
                   Form.getRange('K5').getValue(),    // Produto
                   Form.getRange('K6').getValue(),    // Quantidade
                   Form.getRange('L7').getValue(),    // Custo de Venda
                   Form.getRange('K7').getValue(),    // Preço de Venda                   
                   Form.getRange('K8').getValue()]];  //Total
    
  return Form.getRange(x).setValues(values);
   
};

function inserirProdutoVenda(){

var spreadsheet = SpreadsheetApp.getActive(); 
  
  if(spreadsheet.getRange('AM3').getValue() > 0){
   Browser.msgBox("Erro:","Necessário preencher os campos de produto!",Browser.Buttons);
   spreadsheet.getRange('K4').activate();

  }else{
  
       if (spreadsheet.getRange('G11').getValue() == ""){
           inserirProdutoVenda2('G11:M11');
         
 }else if (spreadsheet.getRange('G12').getValue() == ""){
           inserirProdutoVenda2('G12:M12');
   
 }else if (spreadsheet.getRange('G13').getValue() == ""){
           inserirProdutoVenda2('G13:M13');
    
  }else if (spreadsheet.getRange('G14').getValue() == ""){
            inserirProdutoVenda2('G14:M14');
    
  }else if (spreadsheet.getRange('G15').getValue() == ""){
            inserirProdutoVenda2('G15:M15');
    
  }else if (spreadsheet.getRange('G16').getValue() == ""){
            inserirProdutoVenda2('G16:M16');
    
  }  else {
          Browser.msgBox("Erro:","Todas as linhas foram preenchidas, finalize a venda!",Browser.Buttons);
          }
    
            // Refórmulação
           spreadsheet.getRangeList(['K4:K8']).clear({contentsOnly: true, skipFilteredRows: true});
           spreadsheet.getRange('K7').setFormula('=IF(K6="";"";SWITCH(K5;"P13";AV4;"P20";AW4;"P45";AX4 ))'); //Valor produto
           spreadsheet.getRange('K8').setFormula('=IF(K7="";""; K6*K7)'); //Valor total produto
           spreadsheet.getRange('I25').activate().setFormula('=IF(L18="";"";L18)');
     }
};


//Modo Salvar
function modoNovaVenda() {
  
  var spreadsheet = SpreadsheetApp.getActive(); 
  
  spreadsheet.getRange('AL3').setValue(1);
  spreadsheet.getRange('D1').setValue("Novo");
  //Limpar
  spreadsheet.getRangeList(['G5','H6:H8','K4:K8','K25','G11:M15']).clear({contentsOnly: true, skipFilteredRows: true});
  //Fomrmulação
  spreadsheet.getRange('AN3').setFormula('=IF(G5="";"";QUERY(\'Clientes Dados\'!A:L; "SELECT A, C, L WHERE \'"&G5&"\' = C "))');
  
  spreadsheet.getRange('D4').clearDataValidations().setBackground('#efefef').setFormula('=IF(G5="";"";MAX(\'Vendas Dados\'!A2:A)+1)');
  spreadsheet.getRange('D5').setFormula('=IF(G5="";"";AN4)');
  spreadsheet.getRange('D6').setFormula('=IF(G5="";"";AP4)');

  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('\'Clientes Dados\'!$C$2:$C'), true).build());;
  spreadsheet.getRange('H6').setFormula('=IF(G5="";"";Today())');

  spreadsheet.getRange('K4').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('\'Parâmetros\'!$B$2:$B'), true).build());;
  spreadsheet.getRange('K5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('\'Vendas\'!$BB$4:$BB'), true).build());;  
  spreadsheet.getRange('K7').setFormula('=IF(K6="";"";AW4)');
  spreadsheet.getRange('K8').setFormula('=IF(K7="";""; K6*K7)');

  spreadsheet.getRange('L5').setFormula('=IF(K5="";"";INDEX(AZ4:AZ;MATCH(K5;BB4:BB)))');
  spreadsheet.getRange('L7').setFormula('=IF(K5="";"";INDEX(BC4:BC;MATCH(K5;BB4:BB)))');
  
  spreadsheet.getRange('G18').setFormula('=IF(G11="";"";COUNTA(G11:G16))');
  spreadsheet.getRange('I18').setFormula('=IF(G11="";"";SUM(J11:J16))');
  //spreadsheet.getRange('K18').setFormula('=IF(G11="";"";SUMPRODUCT(J11:J16;K11:K16)');
  spreadsheet.getRange('K18').setFormula('=IF(G11="";"";J11*K11+J12*K12+J13*K13+J14*K14+J15*K15+J16*K16)');
  spreadsheet.getRange('L18').setFormula('=IF(G11="";"";SUM(M11:M16))');
  
  spreadsheet.getRange('H22').setFormula('=IF(G18="";"";MAX(\'Recebimentos Dados\'!A2:A)+1)'); //ID Recebimento
  
  spreadsheet.getRange('G25').setFormula('=IF(G5="";"";H6)');
  spreadsheet.getRange('I25').setFormula('=IF(L18="";"";L18)');
  spreadsheet.getRange('M25').setFormula('=IF(I25="";"";L18-I25)');
  
   spreadsheet.getRange('G5').activate();
  
  
}


//**** SALVAR Venda
// Função auxiliar de salvarVenda
function salvarVenda2(x, idVenda, idCliente, cliente, dataVenda){
  
  var spreadsheet = SpreadsheetApp.getActive();
  var DadosVendas = spreadsheet.getSheetByName('Vendas Dados');
  var Form = spreadsheet.getSheetByName('Vendas');
  
   var values = [[idVenda,                            // ID Venda
                   idCliente,                         // ID Cliente
                   Form.getRange('D6').getValue(),    // Canal de Venda
                   cliente,                           // Cliente
                   dataVenda,                         // Data Venda
                   Form.getRange('H7').getValue(),    // Entregador
                   Form.getRange('H8').getValue(),    // Veiculo
                   Form.getRange(x,7).getValue(),     // ID Produto
                   Form.getRange(x, 8).getValue(),    // Marca
                   Form.getRange(x, 9).getValue(),    // Produto
                   Form.getRange(x, 10).getValue(),   // Quantidade
                   Form.getRange(x, 11).getValue(),   // Custo de Venda
                   Form.getRange(x, 12).getValue(),   // Preço de venda
                   Form.getRange(x, 13).getValue(),   // Total
                ]]; 
    
    
  
  return DadosVendas.getRange(DadosVendas.getLastRow()+1,1,1,14).setValues(values);

};

// Função condutora para salvar venda

function SalvarVenda() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName('Vendas');
  var recebimentosDados = spreadsheet.getSheetByName('Recebimentos Dados');
  var idVenda = Form.getRange('D4').getValue();
  
  if (spreadsheet.getRange('AK3').getValue() > 0 ) 
  {
    Browser.msgBox("Erro", "Necessário preencher todos os campos referente ao cliente, produto e recebimentos! ", Browser.Buttons.OK)
    
  }
  else{
  
        var idCliente = Form.getRange('D5').getValue();         // ID Cliente
        var cliente = Form.getRange('G5').getValue();           // Cliente
        var dataVenda = Form.getRange('H6').getValue();         // Data Venda
        var idRecebimento = Form.getRange('H22').getValue();    // ID Recebimento
        var dataRecebimento = Form.getRange('G25').getValue();  // Data Recebimento
        var totalVenda = Form.getRange('L18').getValue();       // Total da venda
        var formaPagamento =  Form.getRange('K25').getValue();  // Forma Pagamento
        var valorRecebido = Form.getRange('I25').getValue();    // Valor Recebido
        var restante = Form.getRange('M25').getValue();         // Restante
 
        var linhasProd = Form.getRange('AJ3').getValue();       // Quantidade de produtos lançados
        var lin = 11;
    
          for (let i = 1; i <= linhasProd; i++) {
               salvarVenda2(lin, idVenda, idCliente, cliente, dataVenda, idRecebimento, dataRecebimento,
                            formaPagamento, valorRecebido, restante);
               lin++
           }     

 // Salvar Recebimentos
 
                   recebimentosDados.getRange(recebimentosDados.getLastRow()+1,1,1,10).setValues([[idRecebimento, idVenda,
                   idCliente, cliente, dataVenda, dataRecebimento, totalVenda, valorRecebido, formaPagamento, restante]]);

limparVenda();

Browser.msgBox("Informativo", "Registro salvo com sucesso!", Browser.Buttons.OK)




  }
  

 };

//*****************************************

//Modo Deletar

function modoDeletarVenda(){

  var spreadsheet = SpreadsheetApp.getActive(); 
  spreadsheet.getRange('AL3').setValue(3);
  spreadsheet.getRange('D1').setValue("Deletar");
 //Query consulta venda
 // spreadsheet.getRange('BF4').setFormula('=IF(G5="";QUERY(\'Vendas Dados\'!A:S;"SELECT *");IF(D4="";QUERY(\'Vendas Dados\'!A:S;"SELECT * WHERE "&D5&" = B");QUERY(\'Vendas Dados\'!A:S;"SELECT * WHERE "&D5&" = B AND "&D4&" = A")))');
  //Limpar
  spreadsheet.getRangeList(['D4','G5','H6:H8','K4:K8','G11:M15']).clear({contentsOnly: true, skipFilteredRows: true});
  
  spreadsheet.getRange('D4').setBackground('#ffffff').activate().setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false).requireValueInRange(spreadsheet.getRange('\'Vendas\'!$BZ$5:$BZ'), true).build()); //ID Venda
  spreadsheet.getRange('D5').setFormula('=IF(G5="";"";AN4)'); //ID Cliente
  spreadsheet.getRange('D6').setFormula('=IF(D4="";"";BH5)'); //Canal de Venda
  
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('\'Vendas Dados\'!$D$2:$D'), true).build());;
  
  spreadsheet.getRange('H6').setFormula('=IF(D4="";"";BJ5)'); //Data Venda
  spreadsheet.getRange('H7').setFormula('=IF(D4="";"";BK5)'); //Entregador
  spreadsheet.getRange('H8').setFormula('=IF(D4="";"";BL5)'); //Veículo
  
   //Formular aŕea de produto
       
var values = [['=IF($D$4="";"";BM5)','=IF($D$4="";"";BN5)','=IF($D$4="";"";BO5)','=IF($D$4="";"";BP5)','=IF($D$4="";"";BQ5)','=IF($D$4="";"";BR5)','=IF($D$4="";"";BS5)'], // linha 1
              ['=IF($D$4="";"";BM6)','=IF($D$4="";"";BN6)','=IF($D$4="";"";BO6)','=IF($D$4="";"";BP6)','=IF($D$4="";"";BQ6)','=IF($D$4="";"";BR6)','=IF($D$4="";"";BS6)'], // linha 2
              ['=IF($D$4="";"";BM7)','=IF($D$4="";"";BN7)','=IF($D$4="";"";BO7)','=IF($D$4="";"";BP7)','=IF($D$4="";"";BQ7)','=IF($D$4="";"";BR7)','=IF($D$4="";"";BS7)'], // linha 3
              ['=IF($D$4="";"";BM8)','=IF($D$4="";"";BN8)','=IF($D$4="";"";BN8)','=IF($D$4="";"";BP8)','=IF($D$4="";"";BQ8)','=IF($D$4="";"";BR8)','=IF($D$4="";"";BS8)'], // linha 4
              ['=IF($D$4="";"";BM9)','=IF($D$4="";"";BN9)','=IF($D$4="";"";BO9)','=IF($D$4="";"";BP9)','=IF($D$4="";"";BQ9)','=IF($D$4="";"";BR9)','=IF($D$4="";"";BS9)'], // linha 5
              ['=IF($D$4="";"";BM10)','=IF($D$4="";"";BN10)','=IF($D$4="";"";BO10)','=IF($D$4="";"";BP10)','=IF($D$4="";"";BQ10)','=IF($D$4="";"";BR10)','=IF($D$4="";"";BS10)']];  // linha 6

  spreadsheet.getRange('G11:M16').setValues(values);
  
  spreadsheet.getRange('H22').setFormula('=IF(D4="";"";BT5)'); //ID Recebimento
  
  spreadsheet.getRange('G25').setFormula('=IF(D4="";"";BU5)'); //Data Recebimento
  spreadsheet.getRange('I25').setFormula('=IF(D4="";"";BW5)'); //Valor Recebido
  spreadsheet.getRange('K25').setFormula('=IF(D4="";"";BV5)'); //Forma de Pagamento
  spreadsheet.getRange('M25').setFormula('=IF(D4="";"";BX5)'); //Restante
  
  spreadsheet.getRange('G5').activate();

}

//**** Consolidar Deletar Vendas

function DeletarVenda(){
  
  var spreadsheet = SpreadsheetApp.getActive();
  

  if (spreadsheet.getRange('AK3').getValue() > 0 ) {
    
    Browser.msgBox("Erro", "Necessário preencher os campos com ' * ' ", Browser.Buttons.OK);
    
  }else{
  

  var Vendas = spreadsheet.getSheetByName('Vendas');
//var Pesquisa = Vendas.getRange('D4').getValue();
  var VendasDados = spreadsheet.getSheetByName('Vendas Dados');
//var LocalPesquisa = VendasDados.getRange(2, 1, VendasDados.getLastRow()).getValues();
//var Resultado = LocalPesquisa.findIndex(Pesquisa);
//var LINHA = Resultado + 2;
  var LINHA = spreadsheet.getRange('AI3').getValue();
  var quantLinhas = spreadsheet.getRange('AJ3').getValue();
  
  //VendasDados.deleteRow(LINHA);
    
    VendasDados.deleteRows(LINHA, quantLinhas);
    //Limpar
    spreadsheet.getRangeList(['D4','G5','H6:H8','K4:K8','G11:M15']).clear({contentsOnly: true, skipFilteredRows: true});   
    
    Browser.msgBox("Informativo", "Registro deletado!", Browser.Buttons.OK);
  spreadsheet.getRange('H6').setFormula('=IF(D4="";"";BJ5)'); //Data Venda
  spreadsheet.getRange('H7').setFormula('=IF(D4="";"";BK5)'); //Entregador
  spreadsheet.getRange('H8').setFormula('=IF(D4="";"";BL5)'); //Veículo
  
   //Formular aŕea de produto
       
   var values = [['=IF($D$4="";"";BM5)','=IF($D$4="";"";BN5)','=IF($D$4="";"";BO5)','=IF($D$4="";"";BP5)','=IF($D$4="";"";BQ5)','=IF($D$4="";"";BR5)','=IF($D$4="";"";BS5)'],
              ['=IF($D$4="";"";BM6)','=IF($D$4="";"";BN6)','=IF($D$4="";"";BO6)','=IF($D$4="";"";BP6)','=IF($D$4="";"";BQ6)','=IF($D$4="";"";BR6)','=IF($D$4="";"";BS6)'],
              ['=IF($D$4="";"";BM7)','=IF($D$4="";"";BN7)','=IF($D$4="";"";BO7)','=IF($D$4="";"";BP7)','=IF($D$4="";"";BQ7)','=IF($D$4="";"";BR7)','=IF($D$4="";"";BS7)'],
              ['=IF($D$4="";"";BM8)','=IF($D$4="";"";BN8)','=IF($D$4="";"";BN8)','=IF($D$4="";"";BP8)','=IF($D$4="";"";BQ8)','=IF($D$4="";"";BR8)','=IF($D$4="";"";BS8)'],
              ['=IF($D$4="";"";BM9)','=IF($D$4="";"";BN9)','=IF($D$4="";"";BO9)','=IF($D$4="";"";BP9)','=IF($D$4="";"";BQ9)','=IF($D$4="";"";BR9)','=IF($D$4="";"";BS9)'],
              ['=IF($D$4="";"";BM10)','=IF($D$4="";"";BN10)','=IF($D$4="";"";BO10)','=IF($D$4="";"";BP10)','=IF($D$4="";"";BQ10)','=IF($D$4="";"";BR10)','=IF($D$4="";"";BS10)']];

  spreadsheet.getRange('G11:M16').setValues(values);
  
  spreadsheet.getRange('G25').setFormula('=IF(D4="";"";BU5)'); //Data Recebimento
  spreadsheet.getRange('I25').setFormula('=IF(D4="";"";BW5)'); //Valor Recebido
  spreadsheet.getRange('K25').setFormula('=IF(D4="";"";BV5)'); //Forma de Pagamento
  spreadsheet.getRange('M25').setFormula('=IF(D4="";"";BX5)'); //Restante
  
  spreadsheet.getRange('G5').activate();
}
}






//******************    Finalizador   ******************************************************************


function FinalizadorVenda(){

  var spreadsheet = SpreadsheetApp.getActive();

  if(spreadsheet.getRange('AL3').getValue() == 1){
  
    SalvarVenda();
    
  }else if(spreadsheet.getRange('AL3').getValue() == 2){
  
  EditarVenda();
  
  
  }else{
    
    DeletarVenda();  
  
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

function limparVenda(){

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRangeList(['G5','G11:M16','H5:H8','K25']).clear({contentsOnly: true, skipFilteredRows: true}); 

  spreadsheet.getRange('H6').setFormula('=IF(G5="";"";Today())');

  spreadsheet.getRange('K8').setFormula('=IF(K7="";""; K6*K7)');

   spreadsheet.getRange('G5').activate();

}

function limparProdVendas(){

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRangeList(['G11:M16']).clear({contentsOnly: true, skipFilteredRows: true}); 

  spreadsheet.getRange('K7').setFormula('=IF(K6="";"";AW4 )');
  spreadsheet.getRange('K8').setFormula('=IF(K7="";""; K6*K7)');

   spreadsheet.getRange('G5').activate();

}
function relatorioVendas(){

  var spreadsheet = SpreadsheetApp.getActive();
  var url = "https://datastudio.google.com/embed/reporting/2c6f9777-0e85-4943-a07e-9d6482faaa3d/page/Xn1JB"
  var html = "<script> window.open('"+ url + "');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html);
  
  SpreadsheetApp.getUi().showModalDialog(userInterface, "Relatório de Vendas...");
}

