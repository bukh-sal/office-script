
function main(workbook: ExcelScript.Workbook) {
    function getLastRowIdx(inputSheet: ExcelScript.Worksheet, colName: string) {
      let colIdx = 0;
      let lastRow = 2;
      let cellValue = inputSheet.getRange(colName + lastRow).getText();
      while (cellValue != "") {
        let rToCheck = colName + lastRow;
        cellValue = inputSheet.getRange(rToCheck).getText();
        lastRow++;
      }
      lastRow = lastRow - 2;
      return lastRow;
    }
  
    function getLastColumnName(inputSheet: ExcelScript.Worksheet, rowNumber: number) {
      let colIdx = 0;
      let lastCol = "A";
      let cellValue = inputSheet.getRange(lastCol + rowNumber).getText();
      while (cellValue != "") {
        let cToCheck = lastCol + rowNumber;
        cellValue = inputSheet.getRange(cToCheck).getText();
        lastCol = String.fromCharCode(lastCol.charCodeAt(0) + 1);
      }
      lastCol = String.fromCharCode(lastCol.charCodeAt(0) - 2);
      return lastCol;
    }
  
  
    function rowToDict(inputSheet: ExcelScript.Worksheet, headerRow: number, targetRow: number) {
      let lastCol = getLastColumnName(inputSheet, headerRow);
      let dict = {};
      for (let i = 1; i <= lastCol.charCodeAt(0) - 64; i++) {
        let key = inputSheet.getRange(String.fromCharCode(64 + i) + headerRow).getText();
        let val = inputSheet.getRange(String.fromCharCode(64 + i) + targetRow).getText();
        dict[key] = val;
      }
      return dict;
    }
  
  
    let inputSheet = workbook.getWorksheet("Sheet1");
  
    const inCols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];
  
    console.log("last column is " + getLastColumnName(inputSheet, 1));
  
    console.log("last row is " + getLastRowIdx(inputSheet, inCols[0]));
  
    let headerRow = 1;
    let lastRow = getLastRowIdx(inputSheet, inCols[0]);
  
    let dict = rowToDict(inputSheet, headerRow, lastRow);
  
    console.log("col1 = " + dict["col1"]);
    console.log("col2 = " + dict["col2"]);
    console.log("col3 = " + dict["col3"]);
    console.log("col4 = " + dict["col4"]);
    console.log("col5 = " + dict["col5"]);
    console.log("col6 = " + dict["col6"]);
    console.log("col7 = " + dict["col7"]);
    console.log("col8 = " + dict["col8"]);
    console.log("col9 = " + dict["col9"]);
    console.log("col10 = " + dict["col10"]);
  
    // create output sheet
    let outputSheet = workbook.addWorksheet("Sheet2");
    let outCols = ["A", "B", "C", "D", "E"];
    let outColsNames = ["col2", "col4", "col6", "col8", "col10"];
    let outHeaderRow = 1;
    // set row 1
    for (let i = 0; i < outCols.length; i++) {
      outputSheet.getRange(outCols[i] + outHeaderRow).setValue(outColsNames[i]);
    }
  
    // itterate over the rows and fill the output sheet
    for (let i = 2; i <= lastRow; i++) {
      let dict = rowToDict(inputSheet, headerRow, i);
      for (let j = 0; j < outCols.length; j++) {
        outputSheet.getRange(outCols[j] + i).setValue(dict[outColsNames[j]]);
      }
    }
  
    console.log("done");
  
  
  }
  