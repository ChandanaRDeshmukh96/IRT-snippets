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
    document.getElementById("create-table").onclick = createTable; // create table
    document.getElementById("color-cells").onclick = colorCells; // make cells  yellow
    document.getElementById("create-line-chart").onclick = lineChart;
    document.getElementById("compare-charts").onclick = compareCharts;
  }
});

export async function colorCells() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

function createTable(tableName, colNames, tableData, numCol, startCoord, endCoord) {
  Excel.run(function(context) {
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add(startCoord+":"+endCoord, true /*hasHeaders*/);
    var data = tableData?tableData:[
      ["1/1/2017", "The Phone Company", "Communications", "120"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      ["1/11/2017", "Bellows College", "Education", "350.1"],
      ["1/15/2017", "Trey Research", "Other", "135"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ];
    expensesTable.name = tableName? tableName:"ExpensesTable";

    expensesTable.getHeaderRowRange().values = colNames?colNames:[["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, data);

    expensesTable.columns.getItemAt(numCol?numCol:3).getRange().numberFormat = [["€#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();

    return context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function lineChart(){

    var tableName = "ChartTable";
    var colNames = [['Item','Cost']];
    var tableData = [['milk','18'],['sugar','20'],['tea powder','25'],['ginger','40']];
    var numCol = 1;
    var startCoord = 'A10';
    var endCoord = 'B10';
    createTable(tableName, colNames, tableData, numCol, startCoord, endCoord);
    try{
      await Excel.run(async context => {
        var sheet = context.workbook.worksheets.getItem('Sheet1');
        var dataRange = sheet.getRange("A10:B13");
        var chart = sheet.charts.add("Line", dataRange, "auto");
        console.log(sheet.charts);
        chart.title.text = "Sales Data";
        chart.legend.position = "right"
        chart.legend.format.fill.setSolidColor("white");
        chart.dataLabels.format.font.size = 15;
        chart.dataLabels.format.font.color = "black";
        await context.sync();
      }).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
      });
    }catch(err){
      console.log(err);
    }
    
}

async function compareCharts(){

    /**
     * Point to the chart
     * compare the properties of the chart
     * if properties are same, chart can be declared similar
     */

     Excel.run( context =>{
      var sheet = context.workbook.worksheets.getItem("Sheet1" /* sheet name */); 
      /* number of  charts in the sheet */
      var chartCount  = sheet.charts.getCount(); 
      var firstChart = sheet.charts.getItemAt(0 /* index of chart, starting from 0 */);
      var lastChart = sheet.charts.getItemAt(1);

      /** loading all properties that make a difference in the chart */

      firstChart.load(`axes, chartType, dataLabels, format, left, legend, name, pivotOptions, 
      plotBy, series, style, title`);
      firstChart.axes.categoryAxis.title.load('text');
      // console.log(Excel.ChartAxes.getItem('Category', 'Primary'));
      lastChart.load(`axes, chartType, dataLabels, format, left, legend, name, pivotOptions, 
      plotBy, series, style, title`);
      lastChart.series.load('XValues');
      
      return context.sync()
      .then(function () {
      console.log(chartCount.value);

      /** toJSON returns an object containing the properties that were loaded before */
      console.log(firstChart);
      console.log(firstChart.toJSON()); 
      console.log(lastChart.toJSON());

      firstChart.axes.categoryAxis.title.text = 'item';
      });
    }).catch(e => errorHandlerFunction(e))
}

function errorHandlerFunction(e){
  console.log("error occured");
  console.log(e);
}