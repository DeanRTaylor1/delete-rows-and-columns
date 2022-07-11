const ui = SpreadsheetApp.getUi();
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1'); // sheet name
  const rows = sheet.getDataRange().getValues(); // load spreadsheet data into a variable

  let  lastColumn = sheet.getMaxColumns();  
  let  lastRow = sheet.getMaxRows(); 
  console.log(lastColumn);
  console.log(lastRow);




function onOpen() {

  //const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('delete columns & rows');
  menu.addItem('Update spreadsheet Dimensions', 'updateTable');
  menu.addToUi(); 
  
}


let getRow = () => {

  let requiredRows = ui.prompt('Updating Table',
      'How many rows are required?',
      ui.ButtonSet.OK_CANCEL); 
  let button = requiredRows.getSelectedButton();  

    if (button === ui.Button.OK) {

      console.log("the user clicked ok");
      console.log(requiredRows.getResponseText());
      
      let rowNumber = parseInt(requiredRows.getResponseText());
      console.log(isNaN(rowNumber));
      
      
      


        if (isNaN(rowNumber)){              
          
          ui.alert('Please enter a number');
         updateTable();
      }

      else {      
        return rowNumber;   
      }


    }
    else if (button === ui.Button.CLOSE || button === ui.Button.CANCEL){
      console.log("the user clicked close");
      return lastRow;
    }
  

}

 let getCol = () => {

  let requiredColumns = ui.prompt('Updating Table',
      'How many columns are required?',
      ui.ButtonSet.OK_CANCEL);   
  let button = requiredColumns.getSelectedButton();  

    if (button === ui.Button.OK) {

      console.log("the user clicked ok");
      console.log(requiredColumns.getResponseText());
      
      let colNumber = parseInt(requiredColumns.getResponseText());
      console.log(isNaN(colNumber));
      
      
      


        if (isNaN(colNumber)){              
          
          ui.alert('Please enter a number');
         deleteColumns();
      }

      else {      
        return colNumber;   
      }


    }
    else if (button === ui.Button.CLOSE || button === ui.Button.CANCEL){

      console.log("the user clicked close");
      return lastColumn;
    }

  

}

function updateTable () {

  
  let requiredColumns = getCol();
  let requiredRows = getRow();

  //let howManyCol = lastColumn - requiredColumns;


  console.log(lastColumn - requiredColumns);
  console.log(requiredRows);

  
 // console.log('Starting no of rows = ' + rows);
 // console.log('last col no. = ' + lastColumn); 

  if(requiredColumns > lastColumn) {

    let howManyCol = requiredColumns - lastColumn;
    sheet.insertColumns(lastColumn, howManyCol);

  } 

  else { 

    let howManyCol = lastColumn - requiredColumns;
    sheet.deleteColumns(requiredColumns, howManyCol);




  }


  if(requiredRows > lastRow) {

    let howManyRow = requiredRows - lastColumn;
    sheet.insertRows(lastRow, howManyRow);

  } 

  else { 

      let howManyRow = lastRow - requiredRows;
      sheet.deleteRows(requiredRows, howManyRow);

  }
    
    
    

  








