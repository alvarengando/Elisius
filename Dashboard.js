function onOpen() {
   var ui = SpreadsheetApp.getUi();
   
  // Or DocumentApp or FormApp.
  ui.createMenu('Dashboard')
      .addItem('Painel', 'Painel')
      .addSeparator()
     // .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Clientes', 'PagClientes')
          .addSeparator()
          .addItem('Vendas', 'PagVendas')
          .addSeparator()
      .addToUi();



}



function Painel() {
  //SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    // .alert('You clicked the first menu item!');
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Dashboard'), true);
  
}

function PagClientes() {
  var spreadsheet = SpreadsheetApp.getActive();
   spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Clientes'), true);
}




function PagVendas() {
   var spreadsheet = SpreadsheetApp.getActive();
   spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Vendas'), true);
}


function Documentos() {
   var spreadsheet = SpreadsheetApp.getActive();
   spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Documentos'), true);
}

function Vistoria() {
   var spreadsheet = SpreadsheetApp.getActive();
   spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Vistorias'), true);
}


function Vendidos() {
   var spreadsheet = SpreadsheetApp.getActive();
   spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Dados Vendidos'), true);
}
