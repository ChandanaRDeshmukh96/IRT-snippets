// import { isObject } from "util";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("select-toggle-group").onclick = toggleGroup; // Add toggle group 
    document.getElementById("clear-cells").onclick = clearCells;
    
  }
});

function toggleGroup() {
  Excel.run(function (context) {
    // var sheet = context.workbook.worksheets.getActiveWorksheet();
    var selectedRange = context.workbook.getSelectedRange();
    // var usedRange = sheet.getUsedRange();
    selectedRange.load("address");
    var toggleGroupAddress = {};

    return context.sync()
        .then(function () {
            toggleGroupAddress = getToggleGroupAddress(selectedRange.address);
            console.log(`The address of the selected range is "${selectedRange.address}"`);
            console.log(toggleGroupAddress);
            // console.log(`The address of the used range is "${usedRange.address}"`);
        });
}).catch(errorHandlerFunction);
}

function getToggleGroupAddress(address) {
  if (address) {  
    var cellAddress = address.indexOf('!') !== '-1' ? address.split('!')[1] : '';
    var startCell = cellAddress !== '' && cellAddress.indexOf(':') ? cellAddress.split(':')[0] : null;
    var startRow = startCell ? startCell.match(/[\d\.]+|\D+/g)[1] : null;
    var startCol = startCell ? startCell.match(/[\d\.]+|\D+/g)[0] : null;
    var endCell = cellAddress !== '' && cellAddress.indexOf(':') ? cellAddress.split(':')[1] : null;
    var endRow = endCell ? endCell.match(/[\d\.]+|\D+/g)[1] : null;
    var endCol = endCell ? endCell.match(/[\d\.]+|\D+/g)[0] : null;
    return { startCell, endCell, startRow, startCol, endRow, endCol };
  } return {};
}
function clearCells(){
  Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.clear();
    return context.sync();
}).catch(errorHandlerFunction);
}

function errorHandlerFunction(e) {
  console.log("error occured");
  console.log(e);
}
