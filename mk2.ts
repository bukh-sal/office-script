
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
  
    function cleanCommaSeparated(input: string) {
      let asList = input.split(",");
      let cleanList: string[] = [];
      for (let i = 0; i < asList.length; i++) {
        cleanList.push(asList[i].trim());
      }
      // remove duplicates
      let uniqueList = Array.from(new Set(cleanList));
      return uniqueList.join(",");
    }
  
    function getImageName(product_name: string, image_number: number) {
      let cleaned_name = product_name.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase();
      return cleaned_name + "-" + image_number;
    }
  
  
  
    let inputSheet = workbook.getWorksheet("Input");
    // create a new sheet
  
    // if the sheet already exists, delete it
    if (workbook.getWorksheet("Output")) {
      workbook.getWorksheet("Output").delete();
    }
    let outputSheet = workbook.addWorksheet("Output");
    // set the active sheet
    outputSheet.activate()
  
    // test (clone the input sheet)
    let outHeaderCols = ["External ID", "name", "description_ecommerce", "list_price", "cost", "seller_ids/partner_id", "seller_ids/price", "x_studio_brand", "Routes", "Website Product Category / Database ID", "categ_id", "type", "active", "sale_ok", "is_published", "Product Attribute / Attribute", "Product Attribute / Values", "Image", "Extra Product Media/Image", "Extra Product Media/Name"];
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
      if (processedRows > lastRow) {
        done = true;
        break;
      }
  
      let dict = rowToDict(inputSheet, 1, processedRows);
  
      let modedMap = new Map<string, string>();
      modedMap.set("External ID", "__autogen__" + Math.floor(Math.random() * 1000000));
      modedMap.set("name", dict.get("Name"));
      modedMap.set("list_price", dict.get("Sale Price"));
      modedMap.set("cost", dict.get("Cost"));
      modedMap.set("seller_ids/partner_id", dict.get("Vendor"));
      modedMap.set("seller_ids/price", dict.get("Vendor Price"));
      modedMap.set("x_studio_brand", dict.get("Brand"));
      modedMap.set("Routes", "Buy,Dropship");
      modedMap.set("Website Product Category / Database ID", dict.get("Website Category ID"));
      modedMap.set("categ_id", dict.get("Sales Category"));
      modedMap.set("type", "Storable Product");
      modedMap.set("active", "TRUE");
      modedMap.set("sale_ok", "TRUE");
      modedMap.set("is_published", "TRUE");
      modedMap.set("Product Attribute / Attribute", dict.get("Attribute 1 (Name)"));
      modedMap.set("Product Attribute / Values", cleanCommaSeparated(dict.get("Attribute 1 (Values)")));
      modedMap.set("Image", dict.get("Cover Image URL"));
      modedMap.set("Extra Product Media/Image", dict.get("Additional Image URL"));
      modedMap.set("Extra Product Media/Name", getImageName(dict.get("Name"), 1));
      // can be multi line (this causes issue in the output sheet, creating a new row for each line)
      modedMap.set("description_ecommerce", dict.get("Description").replace(/\n/g, " "));
      dictToRow(outputSheet, 1, modedMap, rowToWrite);
      rowToWrite++;
  
      let maxAttr = 5;
      for (let i = 2; i <= maxAttr; i++) {
        if (dict.get("Attribute " + i + " (Name)") != "") {
          let attr = new Map<string, string>();
          attr.set("Product Attribute / Attribute", dict.get("Attribute " + i + " (Name)"));
          attr.set("Product Attribute / Values", cleanCommaSeparated(dict.get("Attribute " + i + " (Values)")));
          dictToRow(outputSheet, 1, attr, rowToWrite);
          rowToWrite++;
        }
      }
  
      processedRows++;
    }
  
  }
  