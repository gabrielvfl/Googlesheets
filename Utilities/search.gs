/*
 Example of using the filter:
 The first argument of the function will be the column you want to use
 The second will be what you want to search for in the column. Remember to enclose the text in quotation marks
 Example: "Insert your text".
 By default, it was defined that an interval that already contains a filter will be flag = 1
 A without filter will be flag = 0.
 Finally, you can enter your search range. A trick that you can use to do
 its range to the last line, is to put the cell address after the ":" without the number
 Example:
 function Button(){
   DoSearch(1,"Something",1,"A8:B");
};
*/

  function DoSearch (Column,Word,Flag,Range){
  var spreadsheet = SpreadsheetApp.getActive();
  var CellValue = SpreadsheetApp.getActive().getRange('A1').getValue();            //Page Name will be write on cell A1
  var Page = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CellValue); 
  Pagina.getActiveCell();
  spreadsheet.getRange(Range).activate();                                          //Range Search
  
  switch(Flag){
    case 0:
        spreadsheet.getActiveSheet().getFilter().remove(); //Case there is a filter, it will be removed
    break;                                                          
    case 1:
      var criteria = SpreadsheetApp.newFilterCriteria()    //It do a search criteria when text is equal to Word. 
      .whenTextEqualTo(Word)                               
      .build()
     break;
   }
    
  spreadsheet.getRange(Range).activate();
  spreadsheet.getRange(Range).createFilter();            //Create a filter in range
  var criteria = SpreadsheetApp.newFilterCriteria()      //It do a search criteria when text is equal to Word. 
  .whenTextEqualTo(Word)                                       
  .build();   
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(Column, criteria);
};
