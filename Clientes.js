/* ********************************************  Inicio Novo Cliente ******************************************* */
//Modo Salvar
function modoNovoCliente() {
  var spreadsheet = SpreadsheetApp.getActive(); 
  
  spreadsheet.getRange('AL3').setValue(1);
  spreadsheet.getRange('D1').setValue("Novo");
  spreadsheet.getRange('D4').setFormula('=IF(G5="";"";MAX(\'Clientes Dados\'!A2:A)+1)');
  spreadsheet.getRange('D5').setFormula('=IF(G5="";"";TODAY())');
  
   spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('\'Clientes Dados\'!$C$2:$C'), false).build());
  
  spreadsheet.getRange('M5').setFormula('=IF(K6="";"";INDEX(BI4:BI;MATCH(K6;BK4:BK)))');
  
  spreadsheet.getRangeList(['G5','G7','G17','H8:H15','K5:K7','J11:M22']).clear({contentsOnly: true, skipFilteredRows: true});
  
};

/* ************************ Salvar Cliente ******************** */

function SalvarCliente() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName('Clientes');
  var clientesDados = spreadsheet.getSheetByName('Clientes Dados');
  
  
  if (spreadsheet.getRange('AK3').getValue() > 0 ) 
  {
    Browser.msgBox("Erro", "Necessário preencher todos os campos: Nome, Logradouro, Município, Bairro e Canal de venda", Browser.Buttons.OK);
  }
  
  else{
    
             
    // Salvar na Página Clientes Dados
                                           
        var values = [[Form.getRange('D4').getValue(),    // ID Cliente
                       Form.getRange('D5').getValue(),    // Data Cadastro
                       Form.getRange('G5').getValue(),    // Nome Cliente
                       Form.getRange('G7').getValue(),    // Logradouro
                       Form.getRange('H8').getValue(),    // Complemento
                       Form.getRange('H9').getValue(),    // Município
                       Form.getRange('H10').getValue(),   // Bairro
                       Form.getRange('H11').getValue(),   // Telefone1
                       Form.getRange('H12').getValue(),   // Telefone2
                       Form.getRange('H13').getValue(),   // Celular1
                       Form.getRange('H14').getValue(),   // Celular2
                       Form.getRange('H15').getValue(),   // Canal de Venda
                       Form.getRange('G17').getValue(),   // Referência
                     ]];
       
          clientesDados.getRange(clientesDados.getLastRow()+1,1,1,13).setValues(values);

                     
                      
           spreadsheet.getRangeList(['G5','G7','G17','H8:H15','K5:K7','J11:M22']).clear({contentsOnly: true, skipFilteredRows: true});
           Browser.msgBox("Informativo", "Registro salvo com sucesso!", Browser.Buttons.OK);
           spreadsheet.getRange('G5').activate();
           } 
         
};

           
                             
//*************************************************************************************************************
/* Inserir Preço produto por cliente*/

function inserirPrecoProdutoCliente2(x){
    
  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName('Clientes');
  
 
  var values = [[Form.getRange('M5').getValue(),    //ID Produto
                 Form.getRange('K5').getValue(),    // Marca
                 Form.getRange('K6').getValue(),    // Produto
                 Form.getRange('K7').getValue(),    // Preço Produto
               ]];  
    
  return Form.getRange(x).setValues(values), spreadsheet.getRangeList(['K5:K7']).clear({contentsOnly: true, skipFilteredRows: true});;
   

};

function inserirPrecoProdutoCliente(){
 
 var spreadsheet = SpreadsheetApp.getActive(); 
 var produtoRep = spreadsheet.getRange('AH3').getValue()
  
  if(spreadsheet.getRange('AM3').getValue() > 0){
   Browser.msgBox("Erro:","Necessário preencher os campos do produto!",Browser.Buttons);
   spreadsheet.getRange('K5').activate();

  }else{
    
  
       if (spreadsheet.getRange('J11').getValue() == "" &&  produtoRep == 0){
           inserirPrecoProdutoCliente2('J11:M11');
         
 }else if (spreadsheet.getRange('J12').getValue() == "" && produtoRep == 0){
           inserirPrecoProdutoCliente2('J12:M12');
   
 }else if (spreadsheet.getRange('J13').getValue() == "" && produtoRep == 0){
           inserirPrecoProdutoCliente2('J13:M13');
    
  }else if (spreadsheet.getRange('J14').getValue() == "" && produtoRep == 0){
            inserirPrecoProdutoCliente2('J14:M14');
    
  }else if (spreadsheet.getRange('J15').getValue() == "" && produtoRep == 0){
            inserirPrecoProdutoCliente2('J15:M15');
    
  }else if (spreadsheet.getRange('J16').getValue() == "" && produtoRep == 0){
            inserirPrecoProdutoCliente2('J16:M16');
    
  }else if (spreadsheet.getRange('J17').getValue() == "" && produtoRep == 0){
            inserirPrecoProdutoCliente2('J17:M17');
    
  }else if (spreadsheet.getRange('J18').getValue() == "" && produtoRep == 0){
            inserirPrecoProdutoCliente2('J18:M18');
    
  }else if (spreadsheet.getRange('J19').getValue() == "" && produtoRep == 0){
            inserirPrecoProdutoCliente2('J19:M19');
    
  }else if (spreadsheet.getRange('J20').getValue() == "" && produtoRep == 0){
            inserirPrecoProdutoCliente2('J20:M20');
    
  }else if (spreadsheet.getRange('J21').getValue() == "" &&produtoRep == 0){
            inserirPrecoProdutoCliente2('J21:M21');
    
  }else if (spreadsheet.getRange('J22').getValue() == "" && produtoRep == 0){
            inserirPrecoProdutoCliente2('J22:M22');
    
  }  else {
                if ( produtoRep > 0){
                      Browser.msgBox("Erro:","Produto já lançado!",Browser.Buttons);
                      spreadsheet.getRange('K5').activate();
                }else{
          Browser.msgBox("Erro:","Todas as linhas foram preenchidas, edite o preço cadastrado",Browser.Buttons);
                     }
           }
       }
};



/* ******************************************** Término Nova Cliente ******************************************* */




/* ******************************************** Inicio Editar Cliente ******************************************* */
//Modo Editar

function modoEditarCliente(){
  
  var spreadsheet = SpreadsheetApp.getActive(); 

  spreadsheet.getRange('AL3').setValue(2);
  spreadsheet.getRange('D1').setValue("Editar");
  
  spreadsheet.getRangeList(['G5','G7','G17','H8:H15','K5:K7']).clear({contentsOnly: true, skipFilteredRows: true});
  
  spreadsheet.getRange('D4').setFormula('=IF(G5="";"";AN4)');
  spreadsheet.getRange('D5').setFormula('=IF(G5="";"";AO4)');
  
   spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('\'Clientes Dados\'!$C$2:$C'), true).build()); //Cliente
  
  spreadsheet.getRange('G7').setFormula('=IF(G5="";"";AQ4)'); // Logradouro
  
  spreadsheet.getRange('H8').setFormula('=IF(G5="";"";AR4)'); // Complemento
  spreadsheet.getRange('H9').setFormula('=IF(G5="";"";AS4)'); // Município
  spreadsheet.getRange('H10').setFormula('=IF(G5="";"";AT4)'); // Bairro
  spreadsheet.getRange('H11').setFormula('=IF(G5="";"";AU4)'); // Telefone 1
  spreadsheet.getRange('H12').setFormula('=IF(G5="";"";AV4)'); // Telefone 2
  spreadsheet.getRange('H13').setFormula('=IF(G5="";"";AW4)'); // Celular 1
  spreadsheet.getRange('H14').setFormula('=IF(G5="";"";AX4)'); // Celular 2
  spreadsheet.getRange('H15').setFormula('=IF(G5="";"";AY4)'); // Canal de Venda
  spreadsheet.getRange('G17').setFormula('=IF(G5="";"";AZ4)'); // Referência
  
  //Formular aŕea de produto
   
  spreadsheet.getRange('J11').setFormula('=IF($G$5="";"";BD4)'); // ID Produto
  spreadsheet.getRange('K11').setFormula('=IF($G$5="";"";BE4)'); // Marca
  spreadsheet.getRange('L11').setFormula('=IF($G$5="";"";BF4)'); // Produto
  spreadsheet.getRange('M11').setFormula('=IF($G$5="";"";BG4)'); // Preço
  
  spreadsheet.getRange('J11:M11').autoFill(spreadsheet.getRange('J11:M22'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  
  spreadsheet.getRange('G5').activate();
};
    
/*  ***********************************************  */    

    //Salvar Editar
   
 function editarCliente(){
    
  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName('Clientes');
  var clientesDados = spreadsheet.getSheetByName('Clientes Dados');
  var linhaCliente = spreadsheet.getRange('AG3').getValue(); //linha correspondente em Clientes dados

   
  if (spreadsheet.getRange('AK3').getValue() > 0 ) 
  {
    Browser.msgBox("Erro", "Necessário preencher todos os campos: Nome, Logradouro, Município, Bairro e Canal de venda", Browser.Buttons.OK);
  }
  
  else{
    
    // Salvar na Página Clientes Dados
                                           
        var values = [[Form.getRange('D4').getValue(),    // ID Cliente
                       Form.getRange('D5').getValue(),    // Data Cadastro
                       Form.getRange('G5').getValue(),    // Nome Cliente
                       Form.getRange('G7').getValue(),    // Logradouro
                       Form.getRange('H8').getValue(),    // Complemento
                       Form.getRange('H9').getValue(),    // Município
                       Form.getRange('H10').getValue(),   // Bairro
                       Form.getRange('H11').getValue(),   // Telefone1
                       Form.getRange('H12').getValue(),   // Telefone2
                       Form.getRange('H13').getValue(),   // Celular1
                       Form.getRange('H14').getValue(),   // Celular2
                       Form.getRange('H15').getValue(),   // Canal de Venda
                       Form.getRange('G17').getValue(),   // Referência
                     ]];
       
          clientesDados.getRange(linhaCliente, 1, 1, 13).setValues(values);
                     
                      
        //   spreadsheet.getRangeList(['G5','G17','H8:H15','K5:K7','G7','J11:M22']).clear({contentsOnly: true, skipFilteredRows: true});
           Browser.msgBox("Informativo", "Registro Alterado com sucesso!", Browser.Buttons.OK);
           spreadsheet.getRange('G7').setFormula('=IF(G5="";"";AQ4)'); // Logradouro
  
           spreadsheet.getRange('H8').setFormula('=IF(G5="";"";AR4)'); // Complemento
           spreadsheet.getRange('H9').setFormula('=IF(G5="";"";AS4)'); // Município
           spreadsheet.getRange('H10').setFormula('=IF(G5="";"";AT4)'); // Bairro
           spreadsheet.getRange('H11').setFormula('=IF(G5="";"";AU4)'); // Telefone 1
           spreadsheet.getRange('H12').setFormula('=IF(G5="";"";AV4)'); // Telefone 2
           spreadsheet.getRange('H13').setFormula('=IF(G5="";"";AW4)'); // Celular 1
           spreadsheet.getRange('H14').setFormula('=IF(G5="";"";AX4)'); // Celular 2
           spreadsheet.getRange('H15').setFormula('=IF(G5="";"";AY4)'); // Canal de Venda
           spreadsheet.getRange('G17').setFormula('=IF(G5="";"";AZ4)'); // Referência
           spreadsheet.getRange('G5').activate();
         
         } 
};  
    
    
    
// Salvar na Página Clientes Preços
function editarPrecosClientes(){
  
  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName('Clientes');
  var clientesPrecos = spreadsheet.getSheetByName('Clientes Preços');
  var linhasProd = spreadsheet.getRange('AJ3').getValue();
  var primeiroProd = spreadsheet.getRange('J11').getValue();
  var t = 11;
  var values = [[],[],[],[],[],[],[],[],[],[],[],[]];
  // var lcp = spreadsheet.getRange('AF3').getValue(); //linha correspondente em Clientes preços
  var idCliente = Form.getRange('D4').getValue();    // ID Cliente, a ser pesquisado
  var localPesquisa = clientesPrecos.getRange(2,1, clientesPrecos.getLastRow()).getValues(); //obtém os valores a serem pesquisados
  var Resultado = localPesquisa.findIndex(idCliente); //linha com o resultado
  var lcp = Resultado +1; // acrescenta 1 devido o cabeçalho 

     if (primeiroProd != "")
      {
         for(var i = 0; i < linhasProd; i++)
            {
              values[i].push(idCliente,                          // ID Cliente
                             Form.getRange('G5').getValue(),     // Nome
                             Form.getRange(t, 10).getValue(),    // ID Produto
                             Form.getRange(t, 11).getValue(),    // Marca
                             Form.getRange(t, 12).getValue(),    // Produto
                             Form.getRange(t, 13).getValue())    // Preço
                        
              t++;
             }   
  
           var values2 = values.slice(0,linhasProd);
        if(lcp <= 0)
          {
            
            clientesPrecos.getRange(clientesPrecos.getLastRow()+1,1,linhasProd,6).setValues(values2);
            //clientesPrecos.getRange('A2:F').sort(1);
            Browser.msgBox("Informativo", "Registro Salvo com sucesso!", Browser.Buttons.OK);
            spreadsheet.getRange('J11').setFormula('=IF($G$5="";"";BD4)'); // ID Produto
            spreadsheet.getRange('K11').setFormula('=IF($G$5="";"";BE4)'); // Marca
            spreadsheet.getRange('L11').setFormula('=IF($G$5="";"";BF4)'); // Produto
            spreadsheet.getRange('M11').setFormula('=IF($G$5="";"";BG4)'); // Preço
            spreadsheet.getRange('J11:M11').autoFill(spreadsheet.getRange('J11:M22'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

          }else{
          
                //  clientesPrecos.getRange(lcp, 1, linhasProd, 6).setValues(values2);
               
                clientesPrecos.deleteRows(lcp, linhasProd); 
                clientesPrecos.getRange(clientesPrecos.getLastRow()+1,1,linhasProd,6).setValues(values2);
               // clientesPrecos.getRange('A2:F').sort(1); 
                Browser.msgBox("Informativo", "Preço salvo com sucesso!", Browser.Buttons.OK);
                spreadsheet.getRange('J11').setFormula('=IF($G$5="";"";BD4)'); // ID Produto
                spreadsheet.getRange('K11').setFormula('=IF($G$5="";"";BE4)'); // Marca
                spreadsheet.getRange('L11').setFormula('=IF($G$5="";"";BF4)'); // Produto
                spreadsheet.getRange('M11').setFormula('=IF($G$5="";"";BG4)'); // Preço
                spreadsheet.getRange('J11:M11').autoFill(spreadsheet.getRange('J11:M22'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
                 
                }   
        
      }
                                 
};

function deletarPrecosClientes(){
                             
  var spreadsheet = SpreadsheetApp.getActive();
  var clientesPrecos = spreadsheet.getSheetByName('Clientes Preços');
  var linhasProd = spreadsheet.getRange('AD3').getValue();
 // var linhasPreco = spreadsheet.getRange('AF3').getValue();                          
  var primeiroProd = spreadsheet.getRange('J11').getValue();

  var idCliente = spreadsheet.getRange('D4').getValue();    // ID Cliente, a ser pesquisado
  var localPesquisa = clientesPrecos.getRange(2,1, clientesPrecos.getLastRow()).getValues(); //obtém os valores a serem pesquisados
  var Resultado = localPesquisa.findIndex(idCliente); //linha com o resultado
  var lcp = Resultado +2; // acrescenta 1 devido o cabeçalho 
Logger.log(lcp,linhasProd);

      if(primeiroProd == "")
        {
           Browser.msgBox("Erro", "Não há registos a serem deletados!", Browser.Buttons.OK);

         }else
              {
                  clientesPrecos.deleteRows(lcp, linhasProd);
                  Browser.msgBox("Informativo", "Registro excluido com sucesso!", Browser.Buttons.OK);          
              }                    
                             
};                             
                  
    
//***************************************** Inicio da funcionalidade de deletar **********************************************



//Modo Deletar


function modoDeletarCliente(){

  var spreadsheet = SpreadsheetApp.getActive(); 
  
  spreadsheet.getRange('AL3').setValue(3);
  spreadsheet.getRange('D1').setValue("Deletar");
  
  spreadsheet.getRangeList(['G5','G7','G17','H8:H15','K5:K7']).clear({contentsOnly: true, skipFilteredRows: true});
  
  spreadsheet.getRange('G5').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true).requireValueInRange(spreadsheet.getRange('\'Clientes Dados\'!$C$2:$C'), true).build()); //Cliente
  

  spreadsheet.getRange('G7').setFormula('=IF(G5="";"";AQ4)'); // Logradouro
  
  spreadsheet.getRange('H8').setFormula('=IF(G5="";"";AR4)'); // Complemento
  spreadsheet.getRange('H9').setFormula('=IF(G5="";"";AS4)'); // Município
  spreadsheet.getRange('H10').setFormula('=IF(G5="";"";AT4)'); // Bairro
  spreadsheet.getRange('H11').setFormula('=IF(G5="";"";AU4)'); // Telefone 1
  spreadsheet.getRange('H12').setFormula('=IF(G5="";"";AV4)'); // Telefone 2
  spreadsheet.getRange('H13').setFormula('=IF(G5="";"";AW4)'); // Celular 1
  spreadsheet.getRange('H14').setFormula('=IF(G5="";"";AX4)'); // Celular 2
  spreadsheet.getRange('H15').setFormula('=IF(G5="";"";AY4)'); // Canal de Venda
  spreadsheet.getRange('G17').setFormula('=IF(G5="";"";AZ4)'); // Referência
  
  //Formular aŕea de produto
   
  spreadsheet.getRange('J11').setFormula('=IF($G$5="";"";BD4)'); // ID Produto
  spreadsheet.getRange('K11').setFormula('=IF($G$5="";"";BE4)'); // Marca
  spreadsheet.getRange('L11').setFormula('=IF($G$5="";"";BF4)'); // Produto
  spreadsheet.getRange('M11').setFormula('=IF($G$5="";"";BG4)'); // Preço
  
  spreadsheet.getRange('J11:M11').autoFill(spreadsheet.getRange('J11:M22'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  
  spreadsheet.getRange('G5').activate();
  
}

    
/*  *************************************************************************  */    

    //Salvar Deletar
   
 function deletarCliente(){
    
  var spreadsheet = SpreadsheetApp.getActive();
  var clientesDados = spreadsheet.getSheetByName('Clientes Dados');
  var clientesPrecos = spreadsheet.getSheetByName('Clientes Preços');
  var linhaCliente = spreadsheet.getRange('AG3').getValue(); //linha correspondente em Clientes dados
  var linhasProd = spreadsheet.getRange('AD3').getValue();   // Quantidade de produtos(linhas) em Clientes Preços
  var linhasPreco = spreadsheet.getRange('AF3').getValue();  // Linha da localização do cliente em Clientes Preços
   
  if (spreadsheet.getRange('AK3').getValue() > 0 ) 
  {
    Browser.msgBox("Erro", "Necessário preencher todos os campos: Nome, Logradouro, Município, Bairro e Canal de venda", Browser.Buttons.OK);
  }
  
  else{
    
        clientesDados.deleteRow(linhaCliente);      // Deletar na Página Clientes Dados
         
        if(linhasProd != 0){
          
           clientesPrecos.deleteRows(linhasPreco, linhasProd);
           }
             
             Browser.msgBox("Informativo", "Registro Deletado com sucesso!", Browser.Buttons.OK);
             modoDeletarCliente();
             
        } 
};  
 
   
   
   
   
   
   
   
   
   
   
   


    
//******************    Finalizador   ******************************************************************


function FinalizadorCliente(){

  var spreadsheet = SpreadsheetApp.getActive();

  if(spreadsheet.getRange('AL3').getValue() == 1)
  {
      SalvarCliente();
  }
                      
  else if(spreadsheet.getRange('AL3').getValue() == 2)
   {
         editarCliente();
   }
                      
   else
   {
    deletarCliente(); 
   } 


};


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

function relatorioClientes(){

  var url = 'https://datastudio.google.com/embed/reporting/8b7a893c-da04-45fa-941f-0910f95d3d86/page/Xn1JB';
  var name = 'Clientes Sintético';
  var url2 = 'https://datastudio.google.com/embed/reporting/d2bf6792-98d5-4c2d-aefc-1c9cf38551d0/page/Xn1JB';
  var name2 = 'Relatório de Preços Clientes';  
  var html = '<html><body><a href="'+url+'" target="blank" onclick="google.script.host.close()">'+name+'</a> <br><br/><a href="'+url2+'" target="blank" onclick="google.script.host.close()">'+name2+'</a></body></html>';
  var ui = HtmlService.createHtmlOutput(html)

  SpreadsheetApp.getUi().showModelessDialog(ui,"Escolha o tipo de Relatório:");
}

