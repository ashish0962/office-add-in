/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
import './taskpane.css'

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
      console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("filter-table").addEventListener('click', filterTable)//.onclick = filterTable;
    document.getElementById("sort-table").onclick = sortTable;
    document.getElementById("create-chart").onclick = createChart;
    document.getElementById("freeze-header").onclick = freezeHeader;
    // document.getElementById("open-dialog").onclick = openDialog;
    document.getElementById("select-option").onchange = doSomething;
  }
});

function createTable() {
  Excel.run(function (context) {

      // TODO1: Queue table creation logic here.
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
      expensesTable.name = "ExpensesTable";

      // TODO2: Queue commands to populate the table with data.
      expensesTable.getHeaderRowRange().values =
          [["Date", "Merchant", "Category", "Amount"]];

      expensesTable.rows.add(null /*add at the end*/, [
          ["1/1/2017", "The Phone Company", "Communications", "120"],
          ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
          ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
          ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
          ["1/11/2017", "Bellows College", "Education", "350.1"],
          ["1/15/2017", "Trey Research", "Other", "135"],
          ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
      ]);

      // TODO3: Queue commands to format the table.
      expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
      expensesTable.getRange().format.autofitColumns();
      expensesTable.getRange().format.autofitRows();
            return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function filterTable() {
  Excel.run(function (context) {

      // TODO1: Queue commands to filter out all expense categories except
      //        Groceries and Education.
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
      var categoryFilter = expensesTable.columns.getItem('Category').filter;
      categoryFilter.applyValuesFilter(['Education', 'Groceries']);

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function sortTable() {
  Excel.run(function(context){
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var sortFields = [
        {
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    return context.sync()
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function createChart() {
  Excel.run(function(context){
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();

    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'Auto');
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "Left";
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 12;
    chart.dataLabels.format.font.color = "red";
    chart.series.getItemAt(0).name = 'Value in \u20AC';
    return context.sync()
  })
  .catch(function(error){
    console.log("Error: " + error);
    if(error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  })
}

function freezeHeader() {
  Excel.run(function (context) {

      // TODO1: Queue commands to keep the header visible when the user scrolls.
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.freezePanes.freezeRows(1);
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

var dialog = null;

function openDialog() {
  // TODO1: Call the Office Common API that opens a dialog
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/popup.html',
    {height: 45, width: 55},

    // TODO2: Add callback parameter
    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);
}

function processMessage(arg) {
  document.getElementById("user-name").innerHTML = arg.message;
  dialog.close();
}

function revokeFilter() {
  Excel.run(function(context){
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet()
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable')
    expensesTable.clearFilters()
    return context.sync();
  })
  .catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}


function doSomething(e) {
  switch(e.target.value){
    case 'filter-table':
      filterTable()
      break;
    case 'sort-table':
      sortTable()
      break;
    case 'create-chart':
      createChart()
      break;
    case 'freeze-header':
      freezeHeader()
      break;
    case 'revoke-filter':
      revokeFilter()
      break;
    default:
      console.log("selected value is " + e.target.value)
  }
}



