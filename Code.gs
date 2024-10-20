/**
 * This function processes the imported data.
 * 
 * @param {Event Object} e : The event object from an installed onChange trigger.
 */
function onChange(e)
{
  try
  {
    processImportedData(e)
  }
  catch (error)
  {
    Logger.log(error['stack'])
    Browser.msgBox(error['stack'])
  }
}

/**
 * This function handles all of the edit events that happen on the spreadsheet, looking out for when the user is trying to use either of the search pages.
 * 
 * @param {Event Object} e : The event object from an installed onEdit trigger.
 */
function installedOnEdit(e)
{
  const range = e.range;
  const col = range.columnStart;
  const row = range.rowStart;
  const rowEnd = range.rowEnd;
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();

  if (sheet.getSheetName() === 'Search for Item Quantity or Amount ($)')
  {
    conditional: if (col == range.columnEnd)
    {
      if (row == 1 && col == 1 && (rowEnd == 16 || rowEnd == 1))
        searchForQuantityOrAmount(spreadsheet, sheet)
      else if (row == rowEnd)
      {
        const numYears = new Date().getFullYear() - 2012 + 1;

        if (row == 2 && col == numYears) // Quantity Data 
          sheet.getRange(3, numYears).uncheck()
        else if (row == 3 && col == numYears) // Amount ($) Data
          sheet.getRange(2, numYears).uncheck()
        else
          break conditional;

        searchForQuantityOrAmount(spreadsheet, sheet)
      }
    }
  }
  else if (sheet.getSheetName() === 'Search for Item Quantites Delimited by Customers')
  {
    if (col == range.columnEnd && row == 1 && col == 1 && (rowEnd == 16 || rowEnd == 1))
      searchForCustomerQuantity(spreadsheet, sheet)
  }
}

// /**
//  * This function creates a new drop-down menu and also deletes the triggers that are not in use.
//  * 
//  * @author Jarren Ralf
//  */
// function installedOnOpen()
// {
//   const ui = SpreadsheetApp.getUi()
//   var triggerFunction;

//   ui.createMenu('PNT Menu')
//     .addSubMenu(ui.createMenu('📑 Display Instructions for Updating Data')
//       .addItem('📉 Invoice', 'display_Invoice_Instructions') 
//       .addItem('📈 Quantity or Amount', 'display_QuantityOrAmount_Instructions'))
//     .addSubMenu(ui.createMenu('📊 Add New Customer')
//       .addItem('🚣‍♂️ Charter or Guide', 'addNewCharterOrGuideCustomer')
//       .addItem('🚢 Lodge', 'addNewLodgeCustomer'))
//     .addSubMenu(ui.createMenu('🖱 Manually Update Data')
//       .addItem('📉 Invoice', 'concatenateAllData')
//       .addItem('📈 Quantity or Amount', 'collectAllHistoricalData'))
//     .addToUi();

//   // Remove all of the unnecessary triggers. When running one-time triggers, they remain attached to the project (but disabled) and the project has a quota of 20 triggers per script
//   ScriptApp.getProjectTriggers().map(trigger => {
//     triggerFunction = trigger.getHandlerFunction();
//     if (triggerFunction != 'onChange' && triggerFunction != 'installedOnEdit' && triggerFunction != 'installedOnOpen') // Keep all of the event triggers
//       ScriptApp.deleteTrigger(trigger)
//   })
// }

/**
 * This function takes all of the yearly data that has been produced and it assimilates it into a set organized by item. Once the data is aggregated,
 * The current inventory, average, and next year prediction is calculated for each item.
 * 
 * @author Jarren Ralf
 */
function concatenateAllCustomerItemData()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const qtySheet = spreadsheet.getSheetByName('Quantity Data')
  const currentYear = new Date().getFullYear();
  const numYears = 3;
  var sheet, sku, items;

  const years = new Array(numYears).fill('').map((_, y) => (currentYear - y).toString());

  const yearsData = years.map(year => {
    spreadsheet.toast('', year, -1)
    sheet = spreadsheet.getSheetByName(year + '_Cust')
    return sheet.getSheetValues(2, 1, sheet.getLastRow() - 1, 4)
  })

  spreadsheet.toast('Yearly data preparation complete. Customer data collection beginning...', '', -1)

  const customerQuantityData = qtySheet.getSheetValues(2, 1, qtySheet.getLastRow() - 1, 1).map(item => {
    item.push('', '', '');
    sku = item[0].split(' - ').pop();

    for (var y = 0; y < numYears; y++)
    {
      items = yearsData[y].filter(itemVal => itemVal[0] === sku);

      if (items)
        item[y + 1] = items.map(val => val[1] + ': [' + val[2] + '] ' + val[3]).join('\n');
    }

    return item;
  })

  spreadsheet.toast('Computations complete. Data being written to spreadsheet...', '', -1)

  const header = ['Item Descriptions', ...years];
  const numRows_AllQty = customerQuantityData.unshift(header)
  const quantityDataSheet = spreadsheet.getSheetByName('Customer Quantity Data');
  quantityDataSheet.clear().getRange(1, 1, numRows_AllQty, customerQuantityData[0].length).setValues(customerQuantityData);
  spreadsheet.getSheetByName('Search for Item Quantites Delimited by Customers').getRange(1, numYears).setValue('Data was last updated on:\n\n' + new Date().toDateString()).offset(0, numYears - 5).activate();
  spreadsheet.toast('All Customer Quantity data has been updated.', 'COMPLETE', -1);
}

/**
 * This function takes all of the yearly invoice data and concatenates it into one meta set of invoice data. This function can be run on its own or
 * it is Trigger via an import of invoice data.
 * 
 * @author Jarren Ralf
 */
function concatenateAllInvoiceData()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const currentYear = new Date().getFullYear()
  var sheet, allData = [];

  new Array(currentYear - 2012 + 1).fill('').map((_, y) => (currentYear - y).toString()).map(year => {
    sheet = spreadsheet.getSheetByName(year)

    if (sheet !== null) // Reverse the data so that it is in descending date (as apposed to ascending), so the concatenations between years is seamless i.e. December 2017 is followed by January 2018
      allData.push(...sheet.getSheetValues(2, 1, sheet.getLastRow() - 1, 8).reverse());
  })

  const lastRow = allData.unshift(['Item Description', 'Customer Name', 'Date', 'Invoice #', 'Location', 'Salesperson', 'Quantity', 'Amount']);
  spreadsheet.getSheetByName('All Data').clearContents().getRange(1, 1, lastRow, 8).setValues(allData)
}

/**
 * This function takes all of the yearly data that has been produced and it assimilates it into a set organized by item. Once the data is aggregated,
 * The current inventory, average, and next year prediction is calculated for each item.
 * 
 * @author Jarren Ralf
 */
function concatenateAllItemData()
{
  const spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.toast('This may take several minutes...', 'Beginning Data Collection', -1)
  const currentYear = new Date().getFullYear();
  const numYears = currentYear - 2012 + 1;
  const years = new Array(numYears).fill('').map((_, y) => (currentYear - y).toString()).reverse(); // Years in ascending order
  const COL = numYears + 4; // A column index to ensure the correct year is being updated when mapping through each year
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
  const itemNum = csvData[0].indexOf('Item #');
  const fullDescription = csvData[0].indexOf('Item List');
  var quanityData = [], amountData = [], sheet, index, year_y;

  // Loop through all of the years
  years.map((year, y) => {
    spreadsheet.toast('', year, -1)
    year_y = COL - y; // The appropriate index for the y-th year

    sheet = spreadsheet.getSheetByName(year)
    sheet.getSheetValues(2, 1, sheet.getLastRow() - 1, 4).map(salesData => { // Loop through all of the sales data for the y-th year
      index = quanityData.findIndex(d => d[0] === salesData[0]); // The index for the current item in the combined quantity data

      if (index !== -1) // Current item is already in combined data list
      {
        quanityData[index][year_y] += Number(salesData[2]) // Increase the quantity
         amountData[index][year_y] += Number(salesData[3]) // Increase the amount ($)
      }
      else // The current item is not in the combined data yet, so add it in
      {
        quanityData.push([salesData[0], salesData[1], 0, 0, 0, ...new Array(numYears).fill(0)])
         amountData.push([salesData[0], salesData[1], 0, 0, 0, ...new Array(numYears).fill(0)])
        quanityData[quanityData.length - 1][year_y] = Number(salesData[2]) // Add quantity to the appropriate year (column)
         amountData[amountData.length  - 1][year_y] = Number(salesData[3]) // Add amount ($) to the appropriate year (column)
      }
    })
  })

  spreadsheet.toast('Calculating averages and predictions...', '', -1)

  var N; // The number of terms in the average
  var n; // The index position used to determine the number of terms in the average
  var totalInventory;

  quanityData = quanityData.map((item, i) => {
    n = 1;

    while (isQtyZero(item[item.length - n]))
      n++;

    N = numYears - n + 1;

    if (N > 1) // Compute the average and make a prediction if we have more than 1 year of data
    {
      [item[3], amountData[i][3]] = getTwoPredictionsUsingLinearRegresssion(
        years.filter((_, y) => y + 1 >= n), // xData
        item.filter((_, t) => t > 4 && t - 5 < N).reverse(),  // yData1
        amountData[i].filter((_, a) => a > 4 && a - 5 < N).reverse(), //yData2
        2025 // X value to predict
      )

      if (item[3] < 0) // If the prediction is negative then we don't want to display it
        item[3] = 0;

      if (amountData[i][3] < 0) // If the prediction is negative then we don't want to display it
        amountData[i][3] = 0;

      for (var r = 5; r < 5 + N; r++)
      {
        item[4] += item[r]; // Quantity Sum
        amountData[i][4] += amountData[i][r]; // Amount Sum
      }
    }
    else // No predictions or averages for only 1 year of data
    {
      item[3] = 0;
      item[4] = 0;
      amountData[i][3] = 0;
      amountData[i][4] = 0;
    }

    adagioInfo = csvData.find(sku => item[0].toString().toUpperCase() == sku[itemNum].toString().toUpperCase());

    if (adagioInfo != null)
    {
      totalInventory = Number(adagioInfo[2]) + Number(adagioInfo[3]) + Number(adagioInfo[4]) + Number(adagioInfo[5]); // adagioInfo[4] is Trites (400) location; Should we add it in??
      item[1] = adagioInfo[fullDescription]; // Update the Adagio description
      item[2] = totalInventory;
      amountData[i][1] = adagioInfo[fullDescription]; // Update the Adagio description
      amountData[i][2] = totalInventory; // Current Inventory
    }
    
    item[4] = Math.round(item[4]*10/N)/10; // Average
    item = item.map(qty => (isQtyNotZero(qty)) ? qty : '') // Remove the zeros, '0', and replace them with a blank string (makes the data present cleaner)

    amountData[i][4] = Math.round(amountData[i][4]*100/N)/100; // Average
    amountData[i] = amountData[i].map(qty => (isQtyNotZero(qty)) ? qty : '') // Remove the zeros

    return item
  })

  spreadsheet.toast('Computations complete. Data being written to spreadsheet...', '', -1)

  const header = ['SKU', 'Descriptions', 'Current Inventory', 'Prediction', 'Average', ...years.reverse()];
  const numRows_AllQty = quanityData.unshift(header)
  const numRows_AllAmt = amountData.unshift(header)
  const quantityDataSheet = spreadsheet.getSheetByName('Quantity Data');
  const   amountDataSheet = spreadsheet.getSheetByName('Amount Data');

  quantityDataSheet.clear().getRange(1, 1, numRows_AllQty, quanityData[0].length).setValues(quanityData);
  amountDataSheet  .clear().getRange(1, 1, numRows_AllAmt,  amountData[0].length).setValues(amountData);
  quantityDataSheet.deleteColumn(1); // Delete SKU column
    amountDataSheet.deleteColumn(1); // Delete SKU column
  spreadsheet.getSheetByName('Search for Item Quantity or Amount ($)').getRange(1, numYears - 2).setValue('Data was last updated on:\n\n' + new Date().toDateString()).offset(0, numYears - 23).activate();
  spreadsheet.toast('All Amount / Quantity data has been updated.', 'COMPLETE', -1)
}

/**
 * This function configures the yearly invoice data into the format that is desired for the spreadsheet to function optimally
 * 
 * @param {Object[][]}    values    : The values of the data that were just imported into the spreadsheet
 * @param {Spreadsheet} spreadsheet : The active spreadsheet
 * @author Jarren Ralf
 */
function configureYearlyInvoiceData(values, spreadsheet)
{/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  const currentYear = new Date().getFullYear();

  values.shift() // Remove the header
  values.pop()   // Remove the final row which contains descriptive stats
  const preData = values.sort(sortByDateThenInvoiveNumber)
  const data = reformatData_YearlyInvoiceData(preData)
  const year = new Array(currentYear - 2012 + 1).fill('').map((_, y) => (currentYear - y).toString()).reverse().find(p => p == data[0][2].getFullYear()) // The year that the data is representing
  const isSingleYear = data.every(date => date[2].getFullYear() == year);

  if (isSingleYear) // Does every line of this spreadsheet contain the same year?
  {
    const numCols = 10;
    const sheets = spreadsheet.getSheets();
    const previousSheet = sheets.find(sheet => sheet.getSheetName() == year)
    var indexAdjustment = 2011

    if (previousSheet)
    {
      indexAdjustment--;
      spreadsheet.deleteSheet(previousSheet)
    }
    
    SpreadsheetApp.flush();
    const newSheet = spreadsheet.insertSheet(year, sheets.length - year + indexAdjustment).hideSheet().setColumnWidths(1, 2, 350).setColumnWidths(3, 7, 85).setColumnWidth(10, 150);
    SpreadsheetApp.flush();
    const lastRow = data.unshift(['Item Description', 'Customer Name', 'Date', 'Invoice #', 'Location', 'Salesperson', 'Quantity', 'Amount', 'Customer', 'Item Number']);
    newSheet.deleteColumns(11, 16)
    newSheet.setTabColor('#a64d79').setFrozenRows(1)
    newSheet.protect()
    newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').offset(0, 0, lastRow, numCols).setNumberFormat('@').setValues(data)

    ScriptApp.newTrigger('concatenateAllInvoiceData').timeBased().after(500).create() // Concatenate all of the data
    spreadsheet.getSheetByName('Search for Invoice #s').getRange(1, 1).activate()
  }
  else
    Browser.msgBox('Incorrect Data', 'Data contains more than one year.', Browser.Buttons.OK)
}

/**
 * This function checks the invoice numbers and reformats the numbers that come from countersales so that they are all displayed in the same format. It also changes
 * the description to the standard Google description so that the items are more easily searched for.
 * 
 * @param {String[][]} preData : The preformatted data.
 * @return {String[][]} The reformatted data
 * @author Jarren Ralf
 */
function reformatData_YearlyInvoiceData(preData)
{
  //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
  const itemNum = csvData[0].indexOf('Item #');
  const fullDescription = csvData[0].indexOf('Item List')
  var item;

  return preData.map(itemVals => {
    item = csvData.find(val => val[itemNum] == itemVals[9])

    if (item != null)
      return (itemVals[3].toString().length === 9 && itemVals[3].toString().charAt('I')) ?
        [item[fullDescription], itemVals[1], itemVals[2], itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]] :
      (itemVals[3].toString().length === 8 && itemVals[3].toString().charAt('I')) ?
        [item[fullDescription], itemVals[1], itemVals[2], '0' + itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]] : 
        [item[fullDescription], itemVals[1], itemVals[2], itemVals[3], itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]]
    else
      return (itemVals[3].toString().length === 9 && itemVals[3].toString().charAt('I')) ? 
        [itemVals[0] + ' - - - - ' + itemVals[9], itemVals[1], itemVals[2], itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]] : 
      (itemVals[3].toString().length === 8 && itemVals[3].toString().charAt('I')) ? 
        [itemVals[0] + ' - - - - ' + itemVals[9], itemVals[1], itemVals[2], '0' + itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]] : 
        [itemVals[0] + ' - - - - ' + itemVals[9], itemVals[1], itemVals[2], itemVals[3], itemVals[4], itemVals[5], itemVals[6], itemVals[7], itemVals[8], itemVals[9]]
  })
}

/**
 * This function takes one set of x values and two sets of y values and it creates a two linear regressions for each set of y values. It takes the 
 * X value for which we want to make a prediction about and it computes at that point.
 * 
 * @param {Number[]} xData  : The data that goes along the x-Axis
 * @param {Number[]} yData1 : The first set of data that goes along the y-Axis
 * @param {Number[]} yData2 : The second set of data that goes along the y-Axis
 * @param {Number}     X    : The x value for which we want to have a prediction about
 * @return {Number[]} Returns the value of the prediction for the first set of y values and the second set, respectively.
 */
function getTwoPredictionsUsingLinearRegresssion(xData, yData1, yData2, X)
{
  var n = xData.length, s2 = 0, cxy1 = 0, cxy2 = 0, cy1 = 0, cy2 = 0;

  var s1 = xData.reduce((total, x_val, i) => {
    s2   += Number(x_val)**2;
    cy1  += Number(yData1[i])
    cy2  += Number(yData2[i])
    cxy1 += Number(yData1[i])*Number(x_val)
    cxy2 += Number(yData2[i])*Number(x_val)
    return total + Number(x_val);
  }, 0)

  const denominator = n*s2 - s1**2;
  const y1 = Math.round(((n*cxy1 - s1*cy1)/denominator*X + (s2*cy1 - s1*cxy1)/denominator)*10)/10;
  const y2 = Math.round(((n*cxy2 - s1*cy2)/denominator*X + (s2*cy2 - s1*cxy2)/denominator)*10)/10;

  return [y1, y2]
}

/**
 * This function checks if a given value is precisely a non-blank string.
 * 
 * @param  {String}  value : A given string.
 * @return {Boolean} Returns a boolean based on whether an inputted string is not-blank or not.
 * @author Jarren Ralf
 */
function isBlank(value)
{
  return value === '';
}

/**
 * This function checks if a given value is precisely not a non-blank string.
 * 
 * @param  {String}  value : A given string.
 * @return {Boolean} Returns a boolean based on whether an inputted string is blank or not.
 * @author Jarren Ralf
 */
function isNotBlank(value)
{
  return value !== '';
}

/**
 * This function checks if a given number is precisely a non-zero number.
 * 
 * @param  {Number}  num : A given number.
 * @return {Boolean} Returns a boolean based on whether an inputted number is not-zero or not.
 * @author Jarren Ralf
 */
function isQtyNotZero(num)
{
  return num !== 0;
}

/**
 * This function checks if a given number is precisely zero.
 * 
 * @param  {Number}  num : A given number.
 * @return {Boolean} Returns a boolean based on whether an inputted number is zero or not.
 * @author Jarren Ralf
 */
function isQtyZero(num)
{
  return num === 0;
}

/**
 * This function process the imported data.
 * 
 * @param {Event Object} : The event object on an spreadsheet edit.
 * @author Jarren Ralf
 */
function processImportedData(e)
{
  if (e.changeType === 'INSERT_GRID')
  {
    var spreadsheet = e.source;
    var sheets = spreadsheet.getSheets();
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3, isYearlyItemData = 4, isYearlyCustomerItemData = 5;

    for (var sheet = sheets.length - 1; sheet >= 0; sheet--) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      if (sheets[sheet].getType() == SpreadsheetApp.SheetType.GRID) // Some sheets in this spreadsheet are OBJECT sheets because they contain full charts
      {
        info = [
          sheets[sheet].getLastRow(),
          sheets[sheet].getLastColumn(),
          sheets[sheet].getMaxRows(),
          sheets[sheet].getMaxColumns(),
          sheets[sheet].getRange(1, 3).getValue().toString().includes('Quantity Specif'), // A characteristic of the item data
          sheets[sheet].getRange(1, 4).getValue().toString().includes('Customer name')    // A characteristic of the customer item data
        ]
      
        // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
        if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || 
            (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0) ||
            info[isYearlyItemData] || info[isYearlyCustomerItemData]) 
        {
          spreadsheet.toast('Processing imported data...', '', 60)
          const values = sheets[sheet].getSheetValues(1, 1, info[numRows], info[numCols]); 
          const sheetName = sheets[sheet].getSheetName()
          const sheetName_Split = sheetName.split(' ')
          const doesPreviousSheetExist = sheetName_Split[1]
          var fileName = sheetName_Split[0]

          if (sheets[sheet].getSheetName().substring(0, 7) !== "Copy Of") // Don't delete the sheets that are duplicates
            spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet that was created

          if (info[isYearlyItemData])
          {
            updateYearlyItemData(values, fileName, doesPreviousSheetExist, spreadsheet)
            concatenateAllItemData()
            spreadsheet.getSheetByName('Search for Item Quantity or Amount ($)').getRange(1, 1).activate()
          }
          else if (info[isYearlyCustomerItemData])
          {
            updateYearlyCustomerItemData(values, fileName, spreadsheet)
            concatenateAllCustomerItemData()
            spreadsheet.getSheetByName('Search for Item Quantites Delimited by Customers').getRange(1, 1).activate()
          }
          
          break;
        }
      }
    }

    // Try and find the file created and delete it
    var file1 = DriveApp.getFilesByName(fileName + '.xlsx')
    var file2 = DriveApp.getFilesByName("Book1.xlsx")

    if (file1.hasNext())
      file1.next().setTrashed(true)

    if (file2.hasNext())
      file2.next().setTrashed(true)
  }
}

/**
 * This function protects all sheets expect for the search pages on the PNT Inventory Prediction Tool spreadsheet, for those, just the relevant cells in the header are protected.
 * 
 * @author Jarren Ralf
 */
function protectAllSheets()
{
  const users = ['triteswarehouse@gmail.com', 'scottnakashima10@gmail.com', 'scottnakashima@hotmail.com', 'pntparksville@gmail.com', 'derykdawg@gmail.com'];
  var sheetName, chartSheet = SpreadsheetApp.SheetType.OBJECT;

  SpreadsheetApp.getActive().getSheets().map(sheet => {
    if (sheet.getType() !== chartSheet)
    {
      sheetName = sheet.getSheetName();

      if (sheetName !== 'Search for Item Quantity or Amount ($)')
      {
        if (sheetName !==  'Search for Invoice #s')
          sheet.protect().addEditor('jarrencralf@gmail.com').removeEditors(users);
        else
          sheet.protect().addEditor('jarrencralf@gmail.com').removeEditors(users).setUnprotectedRanges([sheet.getRange(1, 1, 2)]);
      }
      else
        sheet.protect().addEditor('jarrencralf@gmail.com').removeEditors(users).setUnprotectedRanges([sheet.getRange(1, 1, 3), sheet.getRange(2, 5, 2), sheet.getRange(2, 9, 2), sheet.getRange(3, 11)]);
      }
  })
}

/**
 * This function removes the protections on all sheets.
 * 
 * @author Jarren Ralf
 */
function removeProtectionOnAllSheets()
{
  var chartSheet = SpreadsheetApp.SheetType.OBJECT;

  SpreadsheetApp.getActive().getSheets().map(sheet => {
    if (sheet.getType() !== chartSheet)
      sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0].remove()
  })
}

/**
 * This function searches for either the amount or quantity of product sold to a particular set of customers, 
 * based on which option the user has selected from the checkboxes on the search sheet.
 * 
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf 
 */
function searchForCustomerQuantity(spreadsheet, sheet)
{
  const startTime = new Date().getTime();
  const numCols_SearchSheet = sheet.getLastColumn()
  const searchResultsDisplayRange = sheet.getRange(1, numCols_SearchSheet); // The range that will display the number of items found by the search
  const functionRunTimeRange = sheet.getRange(2, numCols_SearchSheet);      // The range that will display the runtimes for the search and formatting
  const itemSearchFullRange = sheet.getRange(5, 1, sheet.getMaxRows() - 4, numCols_SearchSheet); // The entire range of the Item Search page
  const output = [];
  const searchesOrNot = sheet.getRange(1, 1, 2).clearFormat()                                       // Clear the formatting of the range of the search box
    .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
    .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
    .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
    .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
    .getValue().toString().toUpperCase().split(' NOT ')                                             // Split the search string at the word 'not'

  const searches = searchesOrNot[0].split(' OR ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

  if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
  {
    spreadsheet.toast('Searching...')

    if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
    {
      const dataSheet = spreadsheet.getSheetByName('Customer Quantity Data');
      const data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn());
      const numSearches = searches.length; // The number searches
      var numSearchWords;

      for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
      {
        loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
        {
          numSearchWords = searches[j].length - 1;

          for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
          {
            if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
            {
              if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
              {
                output.push(data[i]);
                break loop;
              }
            }
            else
              break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
          }
        }
      }
    }
    else // The word 'not' was found in the search string
    {
      var dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

      const dataSheet = spreadsheet.getSheetByName('Customer Quantity Data');
      const data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn());
      const numSearches = searches.length; // The number searches
      var numSearchWords;

      for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
      {
        loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
        {
          numSearchWords = searches[j].length - 1;

          for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
          {
            if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
            {
              if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
              {
                for (var l = 0; l < dontIncludeTheseWords.length; l++)
                {
                  if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l]))
                  {
                    if (l === dontIncludeTheseWords.length - 1)
                    {
                      output.push(data[i]);
                      break loop;
                    }
                  }
                  else
                    break;
                }
              }
            }
            else
              break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
          }
        }
      }
    }

    const numItems = output.length;

    if (numItems === 0) // No items were found
    {
      sheet.getRange('A1').activate(); // Move the user back to the seachbox
      itemSearchFullRange.clearContent(); // Clear content
      const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
      const message = SpreadsheetApp.newRichTextValue().setText("No results found.\nPlease try again.").setTextStyle(0, 16, textStyle).build();
      searchResultsDisplayRange.setRichTextValue(message);
    }
    else
    {
      sheet.getRange('A6').activate(); // Move the user to the top of the search items
      itemSearchFullRange.clearContent().setBackground('white'); // Clear content and reset the text format
      sheet.getRange(5, 1, numItems, output[0].length).setNumberFormat('@').setValues(output);
      (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue("1 result found.");
    }

    spreadsheet.toast('Searching Complete.')
  }
  else
  {
    itemSearchFullRange.clearContent(); // Clear content 
    const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
    const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\nPlease try again.").setTextStyle(0, 14, textStyle).build();
    searchResultsDisplayRange.setRichTextValue(message);
  }

  functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * This function searches for either the amount or quantity of product sold to a particular set of customers, 
 * based on which option the user has selected from the checkboxes on the search sheet.
 * 
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf 
 */
function searchForQuantityOrAmount(spreadsheet, sheet)
{
  const startTime = new Date().getTime();
  const numCols_SearchSheet = sheet.getLastColumn()
  const searchResultsDisplayRange = sheet.getRange(1, numCols_SearchSheet - 3); // The range that will display the number of items found by the search
  const functionRunTimeRange = sheet.getRange(2, numCols_SearchSheet - 3);      // The range that will display the runtimes for the search and formatting
  const itemSearchFullRange = sheet.getRange(6, 1, sheet.getMaxRows() - 5, numCols_SearchSheet); // The entire range of the Item Search page
  const checkboxes = sheet.getSheetValues(2, 13, 2, 1);
  const output = [];
  const searchesOrNot = sheet.getRange(1, 1, 3).clearFormat()                                       // Clear the formatting of the range of the search box
    .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
    .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
    .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
    .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
    .getValue().toString().toUpperCase().split(' NOT ')                                             // Split the search string at the word 'not'

  const searches = searchesOrNot[0].split(' OR ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

  if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
  {
    spreadsheet.toast('Searching...')

    if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
    {
      const dataSheet = selectDataSheet(spreadsheet, checkboxes);
      const data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn());
      const numSearches = searches.length; // The number searches
      var numSearchWords;

      for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
      {
        loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
        {
          numSearchWords = searches[j].length - 1;

          for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
          {
            if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
            {
              if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
              {
                output.push(data[i]);
                break loop;
              }
            }
            else
              break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
          }
        }
      }
    }
    else // The word 'not' was found in the search string
    {
      var dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

      const dataSheet = selectDataSheet(spreadsheet, checkboxes);
      const data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn());
      const numSearches = searches.length; // The number searches
      var numSearchWords;

      for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
      {
        loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
        {
          numSearchWords = searches[j].length - 1;

          for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
          {
            if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
            {
              if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
              {
                for (var l = 0; l < dontIncludeTheseWords.length; l++)
                {
                  if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l]))
                  {
                    if (l === dontIncludeTheseWords.length - 1)
                    {
                      output.push(data[i]);
                      break loop;
                    }
                  }
                  else
                    break;
                }
              }
            }
            else
              break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
          }
        }
      }
    }

    const numItems = output.length;

    if (numItems === 0) // No items were found
    {
      sheet.getRange('A1').activate(); // Move the user back to the seachbox
      itemSearchFullRange.clearContent(); // Clear content
      const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
      const message = SpreadsheetApp.newRichTextValue().setText("No results found.\nPlease try again.").setTextStyle(0, 16, textStyle).build();
      searchResultsDisplayRange.setRichTextValue(message);
    }
    else
    {
      var numFormats = (checkboxes[1][0]) ? new Array(numItems).fill(['@', '@', ...new Array(numCols_SearchSheet - 2).fill('$#,##0.00')]) : new Array(numItems).fill([...new Array(numCols_SearchSheet).fill('@')]);
      sheet.getRange('A6').activate(); // Move the user to the top of the search items
      itemSearchFullRange.clearContent().setBackground('white'); // Clear content and reset the text format
      sheet.getRange(6, 1, numItems, output[0].length).setNumberFormats(numFormats).setValues(output);
      (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue("1 result found.");
    }

    spreadsheet.toast('Searching Complete.')
  }
  else
  {
    itemSearchFullRange.clearContent(); // Clear content 
    const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
    const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\nPlease try again.").setTextStyle(0, 14, textStyle).build();
    searchResultsDisplayRange.setRichTextValue(message);
  }

  functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

function moveskutoback()
{
  var splitDescrip;
  const activeRange = SpreadsheetApp.getActiveRange();
  const items = activeRange.getValues().map(item => {
    splitDescrip = item[0].split(' - ');
    splitDescrip.push(splitDescrip.shift())
    return [splitDescrip.join(' - ')]
  })
  activeRange.setValues(items)
}

/**
 * This function returns the sheet that contains the data that the user is interested in. The choice of sheet is directly based on the checkboxes selected on the 
 * item search page.
 * 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @param {Object[][]}  checkboxes  : The values of the checkboxes
 * @author Jarren Ralf 
 */
function selectDataSheet(spreadsheet, checkboxes)
{
  if (checkboxes[0][0]) // Amount
    return spreadsheet.getSheetByName('Quantity Data')
  else if (checkboxes[1][0]) // Quantity
    return spreadsheet.getSheetByName('Amount Data')
}

/**
 * This function take a number and rounds it to two decimals to make it suitable as a price.
 * 
 * @param {Number} num : The given number 
 * @return A number rounded to two decimals
 */
function twoDecimals(num)
{
  return Math.round((num + Number.EPSILON) * 100) / 100
}

/**
 * This function configures the yearly customer item data into the format that is desired for the spreadsheet to function optimally.
 * 
 * @param {Object[][]}      values         : The values of the data that were just imported into the spreadsheet
 * @param {String}         fileName        : The name of the new sheet (which will also happen to be the xlxs file name)
 * @param {Spreadsheet}  spreadsheet       : The active spreadsheet
 * @author Jarren Ralf
 */
function updateYearlyCustomerItemData(values, fileName, spreadsheet)
{
  const currentYear = new Date().getFullYear();
  values.shift() // Remove the header
  values.pop() // Remove the final row
  const yearRange = new Array(currentYear - 2012 + 1).fill('').map((_, y) => (currentYear - y).toString()).reverse()
  var year = yearRange.find(p => p == fileName) // The year that the data is representing

  if (year == null) // The tab name in the spreadsheet doesn't not have the current year saved in it, so the user needs to be prompt so that we know the current year
  {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Enter the year:')

    if (response.getSelectedButton() === ui.Button.OK)
    {
      year = response.getResponseText(); // Text response is assumed to be the year

      if (yearRange.includes(year))
      {
        const numCols = 4;
        const sheets = spreadsheet.getSheets();
        const previousSheet = sheets.find(sheet => sheet.getSheetName() == year + '_Cust')
        var indexAdjustment = 2009

        if (previousSheet)
        {
          indexAdjustment--;
          spreadsheet.deleteSheet(previousSheet)
        }
        
        SpreadsheetApp.flush();
        const newSheet = spreadsheet.insertSheet(year + '_Cust', sheets.length + indexAdjustment - year)
          .setColumnWidth(1, 150).setColumnWidth(2, 75).setColumnWidth(3, 100).setColumnWidth(4, 300);
        SpreadsheetApp.flush();
        const lastRow = values.unshift(['Item Number', 'Quantity', 'Cust #', 'Customer Name']);
        newSheet.deleteColumns(5, 22)
        newSheet.setTabColor('#a64d79').setFrozenRows(1)
        newSheet.protect()
        newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').setBackground('#c0c0c0')
          .offset(0, 0, lastRow, numCols).setHorizontalAlignments(new Array(lastRow).fill(['left', 'right', 'right', 'left'])).setNumberFormat('@').setValues(values)
        newSheet.hideSheet();
      }
      else
      {
        ui.alert('Invalid Input', 'Please import your data again and enter a 4 digit year in the range of [2012, ' + currentYear + '].',)
        return;
      }
    }
    else
    {
      spreadsheet.toast('Data import proccess has been aborted.', '', 60)
      return;
    }
  }
  else
  {
    const numCols = 4;
    const sheets = spreadsheet.getSheets();
    const previousSheet = sheets.find(sheet => sheet.getSheetName() == year + '_Cust')
    var indexAdjustment = 2009

    if (previousSheet)
    {
      indexAdjustment--;
      spreadsheet.deleteSheet(previousSheet)
    }
    
    SpreadsheetApp.flush();
    const newSheet = spreadsheet.insertSheet(year + '_Cust', sheets.length + indexAdjustment - year)
      .setColumnWidth(1, 150).setColumnWidth(2, 75).setColumnWidth(3, 100).setColumnWidth(4, 300);
    SpreadsheetApp.flush();
    const lastRow = values.unshift(['Item Number', 'Quantity', 'Cust #', 'Customer Name']);
    newSheet.deleteColumns(5, 22)
    newSheet.setTabColor('#a64d79').setFrozenRows(1)
    newSheet.protect()
    newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').setBackground('#c0c0c0')
      .offset(0, 0, lastRow, numCols).setHorizontalAlignments(new Array(lastRow).fill(['left', 'right', 'right', 'left'])).setNumberFormat('@').setValues(values)
    newSheet.hideSheet();
  }
}

/**
 * This function configures the yearly customer item data into the format that is desired for the spreadsheet to function optimally.
 * 
 * @param {Object[][]}      values         : The values of the data that were just imported into the spreadsheet
 * @param {String}         fileName        : The name of the new sheet (which will also happen to be the xlxs file name)
 * @param {Boolean} doesPreviousSheetExist : Whether the previous sheet with the same name exists or not
 * @param {Spreadsheet}  spreadsheet       : The active spreadsheet
 * @author Jarren Ralf
 */
function updateYearlyItemData(values, fileName, doesPreviousSheetExist, spreadsheet)
{
  const currentYear = new Date().getFullYear();
  values.shift() // Remove the header
  const yearlyTotalSales = values.pop()[3] // Remove the final row which contains descriptive stats and access the total sales number
  const yearRange = new Array(currentYear - 2012 + 1).fill('').map((_, y) => (currentYear - y).toString()).reverse()
  var year = yearRange.find(p => p == fileName) // The year that the data is representing

  if (year == null) // The tab name in the spreadsheet doesn't not have the current year saved in it, so the user needs to be prompt so that we know the current year
  {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Enter the year:')

    if (response.getSelectedButton() === ui.Button.OK)
    {
      year = response.getResponseText(); // Text response is assumed to be the year

      if (yearRange.includes(year))
      {
        updateYearlySalesData(yearlyTotalSales, year, spreadsheet) // This produces the annual sales chart

        const numCols = 4;
        const sheets = spreadsheet.getSheets();
        const previousSheet = sheets.find(sheet => sheet.getSheetName() == year)
        var indexAdjustment = 2012;

        if (doesPreviousSheetExist || previousSheet)
        {
          indexAdjustment--;
          spreadsheet.deleteSheet(previousSheet)
        }
        
        SpreadsheetApp.flush();
        const newSheet = spreadsheet.insertSheet(year, sheets.length + indexAdjustment - year)
          .setColumnWidth(1, 150).setColumnWidth(2, 700).setColumnWidths(3, 2, 75);
        SpreadsheetApp.flush();
        const lastRow = values.unshift(['Item Number', 'Item Description', 'Quantity', 'Amount']);
        newSheet.deleteColumns(5, 22)
        newSheet.setTabColor('#a64d79').setFrozenRows(1)
        newSheet.protect()
        newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').setBackground('#c0c0c0')
          .offset(0, 0, lastRow, numCols).setHorizontalAlignments(new Array(lastRow).fill(['left', 'left', 'right', 'right'])).setNumberFormat('@').setValues(values)
        newSheet.hideSheet();
      }
      else
      {
        ui.alert('Invalid Input', 'Please import your data again and enter a 4 digit year in the range of [2012, ' + currentYear + '].',)
        return;
      }
    }
    else
    {
      spreadsheet.toast('Data import proccess has been aborted.', '', 60)
      return;
    }
  }
  else
  {
    updateYearlySalesData(yearlyTotalSales, year, spreadsheet) // This produces the annual sales chart

    const numCols = 4;
    const sheets = spreadsheet.getSheets();
    const previousSheet = sheets.find(sheet => sheet.getSheetName() == year)
    var indexAdjustment = 2012

    if (doesPreviousSheetExist)
    {
      indexAdjustment--;
      spreadsheet.deleteSheet(previousSheet)
    }
    
    SpreadsheetApp.flush();
    const newSheet = spreadsheet.insertSheet(year, sheets.length + indexAdjustment - year)
      .setColumnWidth(1, 150).setColumnWidth(2, 700).setColumnWidths(3, 2, 75);
    SpreadsheetApp.flush();
    const lastRow = values.unshift(['Item Number', 'Item Description', 'Quantity', 'Amount']);
    newSheet.deleteColumns(5, 22)
    newSheet.setTabColor('#a64d79').setFrozenRows(1)
    newSheet.protect()
    newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').setBackground('#c0c0c0')
      .offset(0, 0, lastRow, numCols).setHorizontalAlignments(new Array(lastRow).fill(['left', 'left', 'right', 'right'])).setNumberFormat('@').setValues(values)
    newSheet.hideSheet();
  }
}

/**
 * This function takes the yearly total sales from the data that was just imported and it updates the the chart data.
 * 
 * @param   {Number} yearlyTotalSales : The value of the total sales for the particular year.
 * @param   {Number}       year       : The year of the data that has just been imported into the spreadsheet.
 * @param {Spreadsheet} spreadsheet   : The active spreadsheet.
 * @author Jarren Ralf
 */
function updateYearlySalesData(yearlyTotalSales, year, spreadsheet)
{
  const annualSalesDataSheet = spreadsheet.getSheetByName('Annual Sales Data')
  const annualSalesDataRange = annualSalesDataSheet.getRange(4, 1, annualSalesDataSheet.getLastRow() - 3, 2);
  const annualSalesData = annualSalesDataRange.getValues();
  const idx = annualSalesData.findIndex(yr => yr[0] == year);

  if (idx !== -1)
  {
    annualSalesData[idx][1] = yearlyTotalSales;
    annualSalesDataRange.setValues(annualSalesData)
  }
}