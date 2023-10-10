async function main(workbook: ExcelScript.Workbook) {
    // Call the custom REST API with a POST request
    let sheet = workbook.getActiveWorksheet();
    const itemObj : Object = getItem(sheet);
  
    const response = await fetch('https://testservercomp.jcampbell2.repl.co/editItem', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        item: "test",
        type: "QI2 - T-Bushings",
        editItem: itemObj
      })
    });
  
    const data : Object = await response.json();
  
    console.log(data)
  
  }
  
  function getItem(sheet: ExcelScript.Worksheet) : Object {
    const itemListRange : ExcelScript.Range = getRange(sheet.getRange("AR3"));
    const itemObj : Object = getItemListObj(itemListRange, itemListRange.getOffsetRange(0, 3));
    return itemObj;
  }
  
  
  function getItemListObj(propRange: ExcelScript.Range, valueRange: ExcelScript.Range): Object {
    let obj: Object = {};
  
    for (let i = 0; i < propRange.getRowCount(); i++) {
      let prop = propRange.getCell(i, 0).getValue();
      let value = valueRange.getCell(i, 0).getValue();
  
      obj[prop] = value;
    }
  
    return obj;
  }
  
  
  function getRange(startRange: ExcelScript.Range) : ExcelScript.Range {
    let currentCell = startRange;
    let index : number = 0;
    while (currentCell.getValue() != "") {
      currentCell = currentCell.getOffsetRange(1, 0);
      index++;
    }
    let endCell = currentCell.getOffsetRange(-1, 0);
  
    return startRange.getBoundingRect(endCell);
  }