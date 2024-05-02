
function main(workbook: ExcelScript.Workbook) {

    function columnNumberToName(columnNumber: number) {
      let alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
      let columnName = "";
      if (columnNumber <= 26) {
        columnName = alphabet.charAt(columnNumber - 1);
      }
      else {
        let firstLetter = alphabet.charAt(Math.floor((columnNumber - 1) / 26) - 1);
        let secondLetter = alphabet.charAt((columnNumber - 1) % 26);
        columnName = firstLetter + secondLetter;
      }
      return columnName;
    }
  
  
  
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
      let dict = new Map<string, string>();
      for (let i = 1; i <= lastCol.charCodeAt(0) - 64; i++) {
        let key = inputSheet.getRange(String.fromCharCode(64 + i) + headerRow).getText();
        let val = inputSheet.getRange(String.fromCharCode(64 + i) + targetRow).getText();
        dict.set(key, val);
      }
      return dict;
    }
  
    function dictToRow(outputSheet: ExcelScript.Worksheet, headerRow: number, dict: Map<string, string>, tRow: number) {
      let lastCol = getLastColumnName(outputSheet, headerRow);
      let headerCols: string[] = [];
      let headerRange = outputSheet.getRange("A" + headerRow + ":" + lastCol + headerRow);
      let headerValues = headerRange.getTexts();
  
      headerValues[0].forEach((value) => {
        headerCols.push(value);
      });
  
      for (let a = 0; a < dict.size; a++) {
        let key = Array.from(dict.keys())[a];
        let value = Array.from(dict.values())[a];
  
        const result = headerCols
          .map((item, index) => (item === key ? { item, index } : null))
          .filter((x) => x)[0];
  
        let keyLoc: number;
        if (result) {
          keyLoc = result.index;
        } else {
          keyLoc = headerCols.length;
          headerCols.push(key);
          // add column name to the header
          let tCol = columnNumberToName(keyLoc + 1);
          outputSheet.getRange(tCol + headerRow).setValue(key);
        }
        let tCol = columnNumberToName(keyLoc + 1);
        outputSheet.getRange(tCol + tRow).setValue(value);
  
      }
    }
  
  
  
    let inputSheet = workbook.getWorksheet("Sheet1");
    // create a new sheet
  
    // if the sheet already exists, delete it
    if (workbook.getWorksheet("Output")) {
      workbook.getWorksheet("Output").delete();
    }
    let outputSheet = workbook.addWorksheet("Output");
    // set the active sheet
    workbook.setActiveSheet(outputSheet);
  
    // test (clone the input sheet)
    let outHeaderCols = ["External ID", "name", "list_price", "cost", "seller_ids/partner_id", "seller_ids/price", "x_studio_brand", "Routes", "website / category", "active", "sale_ok", "is_published"]
    // add header to output sheet
    for (let i = 0; i < outHeaderCols.length; i++) {
      let col = columnNumberToName(i + 1);
      outputSheet.getRange(col + "1").setValue(outHeaderCols[i]);
    }
  
    let lastRow = getLastRowIdx(inputSheet, "A");
  
    let done = false;
    let rowToWrite = 2
    let processedRows = 2;
  
    while (!done) {
      if (processedRows >= lastRow - 1) {
        done = true;
        break;
      }
  
      let dict = rowToDict(inputSheet, 1, processedRows);
      dict.set("office_script_idx", processedRows.toString());
  
      let modedMap = new Map<string, string>();
      modedMap.set("External ID", "__autogen__" + Math.floor(Math.random() * 1000000));
      modedMap.set("x_studio_brand", dict.get("Brand"));
      modedMap.set("name", dict.get("Name"));
      modedMap.set("list_price", dict.get("Price") * 1.5);
      modedMap.set("cost", dict.get("Cost") * 1.5);
      modedMap.set("seller_ids/partner_id", dict.get("Vendor"));
      modedMap.set("seller_ids/price", dict.get("Vendor Price") * 1.5);
      modedMap.set("office_script_idx", dict.get("office_script_idx"));
      modedMap.set("office_script_idx", "Buy,Dropship");
      modedMap.set("website / category", dict.get("category"));
      modedMap.set("active", "TRUE");
      modedMap.set("sale_ok", "TRUE");
      modedMap.set("is_published", "TRUE");
      dictToRow(outputSheet, 1, modedMap, rowToWrite);
      processedRows++;
      rowToWrite++;
  
  
      let pm_1 = new Map<string, string>();
      pm_1.set("seller_ids/partner_id", "DIY");
      pm_1.set("seller_ids/price", dict.get("Vendor Price") * 2);
      dictToRow(outputSheet, 1, pm_1, rowToWrite);
      rowToWrite++;
  
      let pm_2 = new Map<string, string>();
      pm_2.set("seller_ids/partner_id", "Vendor x");
      pm_2.set("seller_ids/price", dict.get("Vendor Price") * 2.5);
      dictToRow(outputSheet, 1, pm_2, rowToWrite);
      rowToWrite++;
  
    }
  
  }
  