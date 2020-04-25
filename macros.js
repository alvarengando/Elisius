function ASD() {
  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName('Clientes');
  var clientesDados = spreadsheet.getSheetByName('Clientes Dados');
  var clientesPrecos = spreadsheet.getSheetByName('Clientes Preços');
  var linhasProd = spreadsheet.getRange('AJ3').getValue();
  var primeiroProd = spreadsheet.getRange('J11').getValue();
  var linhaCliente = spreadsheet.getRange('AG3').getValue(); //linha correspondente em Clientes dados
  var lcp = spreadsheet.getRange('AF3').getValue(); //linha correspondente em Clientes preços
  var t = 11;

  
  var idCliente = Form.getRange('D4').getValue();
  var nome = Form.getRange('G5').getValue();
  
  var values = [[],[],[],[],[],[],[],[],[],[],[],[]];
  //var values = new Array(linhasProd);
 
    for(var i = 0; i < linhasProd; i++)
     {
  
  values[i].push(Form.getRange('D4').getValue(),    // ID Cliente
                     Form.getRange('G5').getValue(),    // Nome
                     Form.getRange(t, 11).getValue(),    // ID Produto
                     Form.getRange(t, 10).getValue(),    // Marca
                     Form.getRange(t, 12).getValue(),    // Produto
                     Form.getRange(t, 13).getValue())    // Preço
  
       
                        
                          t++;
                           //lcp++;
     }
  
  
  var values2 = values.slice(0,linhasProd);
  clientesPrecos.getRange(lcp, 1, linhasProd, 6).setValues(values2);
  
  /* for(var i = linhasProd +1 ; i < 12; i++)
    {
      delete values[i];
    }
  

  clientesPrecos.getRange(lcp, 1, linhasProd, 6).setValues(values);*/
  Logger.log(values2);
 // Logger.log(values);
  
};
  

function EEAAA() {
  var spreadsheet = SpreadsheetApp.getActive();
  var Form = spreadsheet.getSheetByName('Clientes');
  var clientesDados = spreadsheet.getSheetByName('Clientes Dados');
  var clientesPrecos = spreadsheet.getSheetByName('Clientes Preços');
  var linhasProd = spreadsheet.getRange('AJ3').getValue();
  var primeiroProd = spreadsheet.getRange('J11').getValue();
  var linhaCliente = spreadsheet.getRange('AG3').getValue(); //linha correspondente em Clientes dados
  var lcp = spreadsheet.getRange('AF3').getValue(); //linha correspondente em Clientes preços
  var t = 11;
  
  var values = [] ;
  Logger.log(values);
  
};


function aRRASTAR() {
  var spreadsheet = SpreadsheetApp.getActive();
  //spreadsheet.getRange('J11:M11').activate();
   spreadsheet.getRange('J11:M11').autoFill(spreadsheet.getRange('J11:M22'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  //spreadsheet.getRange('J11:M22').activate();
};

function lastrowlastcoll() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('J11').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Clientes Preços'), true);
  spreadsheet.getRange('A2').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getRange('A16').activate();
  spreadsheet.getRange('Clientes!J11:M13').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};

function searchString(){
  var sheet = SpreadsheetApp.getActiveSheet()
  var search_string = 1  //sheet.getRange('I1').getValue();
  var textFinder = sheet.createTextFinder(search_string)
  var search_row = textFinder.findNext().getRow()
  var ui = SpreadsheetApp.getUi();
  ui.alert("search row: " + search_row)
}

function classificar() {
  var spreadsheet = SpreadsheetApp.getActive();
  var clientesPrecos = spreadsheet.getSheetByName('Clientes Preços');
  clientesPrecos.getRange('A2:F').sort(1);
};

function testeC() {
  var spreadsheet = SpreadsheetApp.getActive();
  
  
  
  var linhasProd = spreadsheet.getRange('G11:G16').getLastRow();
  
  Logger.log( linhasProd);
};