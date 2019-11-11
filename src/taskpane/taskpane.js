// import { isObject } from "util";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

var chartData;
var chartObj;

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
    document.getElementById("create-line-chart").onclick = createChart;
    document.getElementById("get-chart-data").onclick = getChart;
    document.getElementById("recreate-chart").onclick = recreateChart;
    document.getElementById("compare-chart").onclick = compareCharts
    
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

    tableName = typeof(tableName)!== 'string' ? null : tableName;
    /** Plot a table in current sheet */
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

    /** startCoord and endCoord may be graded cells */
    startCoord = startCoord? startCoord : "A1";
    endCoord = endCoord? endCoord : "D1"
    var expensesTable = currentWorksheet.tables.add(startCoord + ":" + endCoord, true /*hasHeaders*/);

    /** Either draw a default table or table for the data passed by other function */

    /** Column names or Headers */
    expensesTable.getHeaderRowRange().values = colNames ? colNames : [["Date", "Merchant", "Category", "Amount"]];
    var data = tableData
      ? tableData
      : [
          ["1/1/2017", "The Phone Company", "Communications", "120"],
          ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
          ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
          ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
          ["1/11/2017", "Bellows College", "Education", "350.1"],
          ["1/15/2017", "Trey Research", "Other", "135"],
          ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
        ];

    /** Table format options */
    tableName = typeof(tableName)!== 'string' ? null : tableName;
    expensesTable.name = tableName ? tableName : "ExpensesTable";
    expensesTable.rows.add(null /*add at the end*/, data);
    expensesTable.columns.getItemAt(numCol ? numCol : 3).getRange().numberFormat = [["€#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();

    /** Sync updates to Excel online */
    return context.sync();

  }).catch(function(error) {
    console.log("Error: " + error);
    // eslint-disable-next-line no-undef
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function createChart() {

  //Create table required to plot a chart
  var tableName = "ChartTable";
  var colNames = [["Item", "Cost"]]; // Header
  var tableData = [["milk", "18"], ["sugar", "20"], ["tea powder", "25"], ["ginger", "40"]]; // Data
  var numCol = 1; // Column index to format table data
  var startCoord = "A1";
  var endCoord = "B1";
  createTable(tableName, colNames, tableData, numCol, startCoord, endCoord);
  
  try {
    await Excel.run(async context => {

      var sheet = context.workbook.worksheets.getItem("Sheet1");

      //Get data range from the table or specify directly in an array
      var dataRange = sheet.tables.getItem(tableName).getRange();

      //Create a line chart
      var chart = sheet.charts.add("Line", dataRange, "auto");

      //Chart format options
      chart.title.text = "Sales Data";
      chart.legend.position = "right";
      chart.legend.format.fill.setSolidColor("white");
      chart.dataLabels.format.font.size = 15;
      chart.dataLabels.format.font.color = "black";

      // position can be mapped to the graded cells.
      chart.setPosition("D1", "H10");

      // Sync all of the changes to Online Excel
      await context.sync();


    }).catch(function(error) {
      console.log("Error: " + error);
      // eslint-disable-next-line no-undef
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  } catch (err) {
    console.log(err);
  }
}

async function getChart(e, chartIndex) {
  /**
   * Point to the chart
   * compare the properties of the chart
   * if properties are same, chart can be declared similar
   */

  Excel.run(context => {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    // var sheet = context.workbook.worksheets.getItem("Sheet1" /* sheet name */);
    /* number of  charts in the sheet */
    sheet.tables.load("items");
    var chart = sheet.charts.getItemAt(chartIndex ? chartIndex: 0 /* index of chart, starting from 0 */);
    var allPropsOfChart = `axes, axis, categoryLabelLevel, chartType, context, dataLabels, format, height, id, left, legend, name, pivotOptions, plotBy, plotArea, seriesNameLevel, showAllFieldButtons, series, style, title, top, width, worksheet`;

    chart.load(allPropsOfChart);
    // chart.axes.categoryAxis.title.load('text,position');
    // chart.axes.valueAxis.title.load('text,position');
    // chart.axes.seriesAxis.title.load('text,position');
    chart.title.load("text,position");  
    chart.legend.format.load('border,font,fill');
    chart.dataLabels.format.font.load("size,color");
    let dataInChart = [];
    var seriesCollection = chart.series;
    // loadAllProps(chart);

    seriesCollection.load("items");
    return context.sync().then(async() => {
      for (var i = 0; i < seriesCollection.items.length; i++) {
        var collectionName = seriesCollection.getItemAt(i);
        collectionName.load("points");
        console.log("Chart Data");
        var categories = collectionName.getDimensionValues("Categories");
        await context.sync().then(()=>{
          dataInChart[i] = [];
        for (var j = 0; j < collectionName.points.items.length; j++) {
          dataInChart[i][j] = [];
          dataInChart[i][j][0] = categories.value[j];
          dataInChart[i][j][1] = collectionName.points.items[j].value;
        }
        });
      }
      chartData = dataInChart;
      console.log(dataInChart);
      chartObj = chart.toJSON();
      console.log("Chart Obj");
      console.log(chart.toJSON())
      return chartObj;
    });
  }).catch(e => errorHandlerFunction(e));
}

async function recreateChart(){

    /** Create table required to plot a chart*/
    var tableName = "ChartTable1";
    var colNames = [["Item", "Cost"]]; // Header
    var tableData = chartData[0]; // Data
    var numCol = 1; // Column index to format table data
    var startCoord = "A12";
    var endCoord = "B12";
    createTable(tableName, colNames, tableData, numCol, startCoord, endCoord);
    
    try {
      await Excel.run(async context => {
  
        var sheet = context.workbook.worksheets.getItem("Sheet1");
  
        /**Get data range from the table or specify directly in an array*/
        var dataRange = sheet.tables.getItem(tableName).getRange();
  
        /**Create a line chart */
        var chart = sheet.charts.add(chartObj.chartType, dataRange, "auto");
  
        /** Chart format options */

        // chart.plotArea.set(chartObj);
        chart.title.text = chartObj.title.text ? chartObj.title.text : '';
        chart.legend.visible = chartObj.legend.visible;
        chart.legend.position = chartObj.legend.position;
        chart.axes.categoryAxis.title.text = chartObj.axes.categoryAxis ? chartObj.axes.categoryAxis.title.text:"";
        chart.axes.valueAxis.title.text = chartObj.axes.valueAxis ? chartObj.axes.valueAxis.title.text : "";
        chart.dataLabels.format.font.size = chartObj.dataLabels.format.font.size ? chartObj.dataLabels.format.font.size : '';
        chart.dataLabels.format.font.color = chartObj.dataLabels.format.font.size ? chartObj.dataLabels.format.font.size : '';
  
        /** Position can be mapped to the graded cells. */ 
        chart.setPosition("D12", "H22");
        /** Sync all of the changes to Online Excel */ 
        await context.sync();
  
  
      }).catch(function(error) {
        console.log("Error: " + error);
        // eslint-disable-next-line no-undef
        if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
      });
    } catch (err) {
      console.log(err);
  }
}

async function compareCharts(){
  Excel.run(async context => {
  var sheet = context.workbook.worksheets.getActiveWorksheet();
    // var sheet = context.workbook.worksheets.getItem("Sheet1" /* sheet name */);
    /* number of  charts in the sheet */
    sheet.charts.load("items");
    var chart1 = sheet.charts.getItemAt(0);
    var chart2 = sheet.charts.getItemAt(1);
    var chart1Data = await getChart(chart1);
    var chart2Data = await getChart(chart2);
    console.log( "chart1Data == chart2Data" + chart1Data === chart2Data );
    return context.sync();
  });
}


function errorHandlerFunction(e) {
  console.log("error occured");
  console.log(e);
}
