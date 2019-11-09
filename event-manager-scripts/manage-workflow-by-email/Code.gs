

//----------------------------
// Post-treatment
//----------------------------
function publicationProcess() {

  clearFilter();
  Logger.log("Filter reseted at: "+ new Date());
  //STEP 1 : Sub process of SUBMITTED ROW (Go to Published state or validated state)
  setFilter(columnIndexForState, criterias[0])
  Logger.log("Filter activated: Submitted data"+ new Date());

  var ss = SpreadsheetApp.SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var lastColumn = sheet.getLastColumn();


  var submittedListOfIndex = getIndexesOfFilteredRows();
  var submittedLengthList = submittedListOfIndex.length;
  if (submittedLengthList > 0) {
    // LOGICAL TESTS
    for (var j = 0; j < submittedLengthList; j++) {
      var currentLigne = submittedListOfIndex[j] + 1;
      var row = proceedByRow(currentLigne, sheet, lastColumn);
      row = rowToObject(row[0]);
      Logger.log(row);
      
 clearFilter();
}
}
}
