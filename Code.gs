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
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();
  const sheetName = sheet.getSheetName();

  try
  {
    if (sheetName === 'Search for Item Quantity or Amount ($)')
      searchForQuantityOrAmount(e, spreadsheet, sheet)
    else if (sheetName === 'Search for Item Quantity partitioned by Customers')
      searchForCustomerQuantity(e, spreadsheet, sheet)
    else if (sheetName === 'Search for Invoice #s')
      searchForInvoice(e, spreadsheet, sheet);
  }
  catch (err)
  {
    var error = err['stack'];
    Logger.log(error)
    Browser.msgBox(error)
    throw new Error(error);
  }
}

/**
 * This function creates a new drop-down menu and also deletes the triggers that are not in use.
 * 
 * @author Jarren Ralf
 */
function installedOnOpen()
{
  var triggerFunction;

  // ui.createMenu('PNT Menu')
  //   .addSubMenu(ui.createMenu('📑 Display Instructions for Updating Data')
  //     .addItem('📉 Invoice', 'display_Invoice_Instructions') 
  //     .addItem('📈 Quantity or Amount', 'display_QuantityOrAmount_Instructions'))
  //   .addSubMenu(ui.createMenu('📊 Add New Customer')
  //     .addItem('🚣‍♂️ Charter or Guide', 'addNewCharterOrGuideCustomer')
  //     .addItem('🚢 Lodge', 'addNewLodgeCustomer'))
  //   .addSubMenu(ui.createMenu('🖱 Manually Update Data')
  //     .addItem('📉 Invoice', 'concatenateAllData')
  //     .addItem('📈 Quantity or Amount', 'collectAllHistoricalData'))
  //   .addToUi();

  // Remove all of the unnecessary triggers. When running one-time triggers, they remain attached to the project (but disabled) and the project has a quota of 20 triggers per script
  ScriptApp.getProjectTriggers().map(trigger => {
    triggerFunction = trigger.getHandlerFunction();
    if (triggerFunction != 'onChange' && triggerFunction != 'installedOnEdit' && triggerFunction != 'installedOnOpen') // Keep all of the event triggers
      ScriptApp.deleteTrigger(trigger)
  })
}

/**
 * This function takes all of the yearly data that has been produced and it assimilates it into a set organized by item. Once the data is aggregated,
 * The current inventory, average, and next year prediction is calculated for each item.
 * 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function concatenateAllCustomerItemData(spreadsheet)
{
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
  spreadsheet.getSheetByName('Search for Item Quantity partitioned by Customers').getRange(1, numYears).setValue('Data was last updated on:\n\n' + new Date().toDateString()).offset(0, numYears - 5).activate();
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

  new Array(3).fill('').map((_, y) => (currentYear - y).toString()).map(year => {
    sheet = spreadsheet.getSheetByName(year + '_Inv')
    spreadsheet.toast('Update complete. Concatenating all invoice data...', year, -1)

    if (sheet !== null) // Reverse the data so that it is in descending date (as opposed to ascending), so the concatenations between years is seamless i.e. December 2017 is followed by January 2018
      allData.push(...sheet.getSheetValues(2, 1, sheet.getLastRow() - 1, 8).reverse());
  })

  spreadsheet.toast('Concatenation complete. Writing data to spreadsheet', '', -1)
  const lastRow = allData.unshift(['Item Description', 'Customer Name', 'Date', 'Invoice #', 'Location', 'Salesperson', 'Quantity', 'Amount']);
  spreadsheet.getSheetByName('Invoice Data').clearContents().getRange(1, 1, 1, 8).setFontSize(11).setFontWeight('bold').setBackground('#c0c0c0').offset(0, 0, lastRow, 8).setValues(allData);
  spreadsheet.toast('Invoice data has been updated', 'COMPLETE', 60)
}

/**
 * This function takes all of the yearly data that has been produced and it assimilates it into a set organized by item. Once the data is aggregated,
 * The current inventory, average, and next year prediction is calculated for each item.
 * 
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function concatenateAllItemData(spreadsheet)
{
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
{
  spreadsheet.toast('Sorting and reformatting the invoice data...', '', -1)
  const currentYear = new Date().getFullYear();
  values.shift() // Remove the header
  values.pop()   // Remove the final row which contains descriptive stats
  const data = reformatData_YearlyInvoiceData(values.sort(sortByDateThenInvoiceNumber))
  const year = new Array(currentYear - 2012 + 1).fill('').map((_, y) => (currentYear - y).toString()).reverse().find(p => p == data[0][2].getFullYear()) // The year that the data is representing
  const isSingleYear = data.every(date => date[2].getFullYear() == year);

  if (isSingleYear) // Does every line of this spreadsheet contain the same year?
  {
    spreadsheet.toast('Sorting and reformatting complete. Updating data for current year...', year, -1)
    const numCols = 8;
    const sheets = spreadsheet.getSheets();
    const previousSheet = sheets.find(sheet => sheet.getSheetName() == year + '_Inv')
    var indexAdjustment = 2009

    if (previousSheet)
    {
      indexAdjustment--;
      spreadsheet.deleteSheet(previousSheet)
    }
    
    SpreadsheetApp.flush();
    const newSheet = spreadsheet.insertSheet(year + '_Inv', sheets.length - year + indexAdjustment).hideSheet().setColumnWidth(1, 800).setColumnWidth(2, 350).setColumnWidths(3, 6, 85);
    SpreadsheetApp.flush();

    const lastRow = data.unshift(['Item Description', 'Customer Name', 'Date', 'Invoice #', 'Location', 'Salesperson', 'Quantity', 'Amount']);
    newSheet.deleteColumns(9, 18)
    newSheet.setTabColor('#a64d79').setFrozenRows(1)
    newSheet.protect()
    newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').setBackground('#c0c0c0').offset(0, 0, lastRow, numCols).setNumberFormat('@').setValues(data)

    ScriptApp.newTrigger('concatenateAllInvoiceData').timeBased().after(500).create() // Concatenate all of the data
    spreadsheet.getSheetByName('Search for Invoice #s').getRange(1, 1).activate();
    spreadsheet.toast('The data will be updated in less than 5 minutes.', 'COMPLETE')
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
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
  const itemNum = csvData[0].indexOf('Item #');
  const fullDescription = csvData[0].indexOf('Item List')
  var item;

  return preData.map(itemVals => {
    item = csvData.find(val => val[itemNum] == itemVals[9])

    if (item != null)
      return (itemVals[3].toString().length === 9 && itemVals[3].toString().charAt('I')) ?
        [item[fullDescription], itemVals[1] + ' [' + itemVals[8] + ']', itemVals[2], itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7]] :
      (itemVals[3].toString().length === 8 && itemVals[3].toString().charAt('I')) ?
        [item[fullDescription], itemVals[1] + ' [' + itemVals[8] + ']', itemVals[2], '0' + itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7]] : 
        [item[fullDescription], itemVals[1] + ' [' + itemVals[8] + ']', itemVals[2], itemVals[3], itemVals[4], itemVals[5], itemVals[6], itemVals[7]]
    else
      return (itemVals[3].toString().length === 9 && itemVals[3].toString().charAt('I')) ? 
        [itemVals[0] + ' - - - - ' + itemVals[9], itemVals[1] + ' [' + itemVals[8] + ']', itemVals[2], itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7]] : 
      (itemVals[3].toString().length === 8 && itemVals[3].toString().charAt('I')) ? 
        [itemVals[0] + ' - - - - ' + itemVals[9], itemVals[1] + ' [' + itemVals[8] + ']', itemVals[2], '0' + itemVals[3].substring(1), itemVals[4], itemVals[5], itemVals[6], itemVals[7]] : 
        [itemVals[0] + ' - - - - ' + itemVals[9], itemVals[1] + ' [' + itemVals[8] + ']', itemVals[2], itemVals[3], itemVals[4], itemVals[5], itemVals[6], itemVals[7]]
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
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3, isYearlyItemData = 4, isYearlyCustomerItemData = 5, isYearlyInvoiceData = 6;

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
          sheets[sheet].getRange(1, 4).getValue().toString().includes('Customer name'),   // A characteristic of the customer item data
          sheets[sheet].getRange(1, 7).getValue().toString().includes('Quantity Specif')  // A characteristic of the invoice data
        ]
      
        // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
        if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || 
            (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0) ||
            info[isYearlyItemData] || info[isYearlyCustomerItemData] || info[isYearlyInvoiceData]) 
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
            concatenateAllItemData(spreadsheet);
          }
          else if (info[isYearlyCustomerItemData])
          {
            updateYearlyCustomerItemData(values, fileName, spreadsheet)
            concatenateAllCustomerItemData(spreadsheet)  
          }
          else if (info[isYearlyInvoiceData])
          {
            configureYearlyInvoiceData(values, spreadsheet)
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
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf 
 */
function searchForCustomerQuantity(e, spreadsheet, sheet)
{
  const startTime = new Date().getTime();
  const range = e.range;
  const row = range.rowStart;
  const rowEnd = range.rowEnd;

  if (row == rowEnd) // Check and make sure only a single row is being edited //  (rowEnd == row || ***rowEnd == 2***)??
  {
    const col = range.columnStart;
    const colEnd = range.columnEnd;

    if ((col == colEnd || colEnd == null) && row == 1 && col == 1) // The search box is being editted
    {
      spreadsheet.toast('Searching...', '', -1)
  
      const numCols_SearchSheet = sheet.getLastColumn()
      var output = [];
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
        if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
        {
          const dataSheet = spreadsheet.getSheetByName('Customer Quantity Data');

          if (searches[0][0] !== 'YEAR') // Check for the year search indicator
          {
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
          else // The year search indicator seems to be found, now check the year
          {
            const currentYear = new Date().getFullYear();
            var y; // The index of the year that the customer is intending to search for
          
            switch (searches[0][1]) // Based on the search indicator, set the column of data to search in
            {
              case currentYear.toString():
                y = 1;
                break;
              case (currentYear - 1).toString():
                y = 2;
                break;
              case (currentYear - 2).toString():
                y = 3;
                break;
            }

            output = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).filter(year => isNotBlank(year[y]));
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
          sheet.getRange(5, 1, sheet.getMaxRows() - 4, numCols_SearchSheet).clearContent(); // Clear content
          const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
          const message = SpreadsheetApp.newRichTextValue().setText("No results found.\nPlease try again.").setTextStyle(0, 16, textStyle).build();
          sheet.getRange(1, numCols_SearchSheet).setRichTextValue(message);
        }
        else
        {
          sheet.getRange('A5').activate(); // Move the user to the top of the search items
          sheet.getRange(5, 1, sheet.getMaxRows() - 4, numCols_SearchSheet).clearContent().setBackground('white'); // Clear content and reset the text format and background colour
          sheet.getRange(5, 1, numItems, numCols_SearchSheet).setNumberFormat('@').setHorizontalAlignments(horizontalAlignments).setValues(output);
          (numItems !== 1) ? sheet.getRange(1, numCols_SearchSheet).setValue(numItems + " results found.") : sheet.getRange(1, numCols_SearchSheet).setValue("1 result found.");
        }

        spreadsheet.toast('Searching Complete.')
      }
      else
      {
        sheet.getRange('A1').activate(); // Move the user back to the seachbox
        sheet.getRange(5, 1, sheet.getMaxRows() - 4, numCols_SearchSheet).clearContent(); // Clear content 
        const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
        const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\nPlease try again.").setTextStyle(0, 14, textStyle).build();
        sheet.getRange(1, numCols_SearchSheet).setRichTextValue(message);
      }

      sheet.getRange(2, numCols_SearchSheet).setValue((new Date().getTime() - startTime)/1000 + " seconds");
    }
  }
  else if (row > 4) // Multiple rows are being edited
  {
    spreadsheet.toast('Searching...', '', -1)
    const numCols_SearchSheet = sheet.getLastColumn()
    sheet.getRange(1, 1, 2).clearContent() // Clear the content for the range of the search box
    const values = range.getValues().filter(blank => isNotBlank(blank[0]))

    if (values.length !== 0) // Don't run function if every value is blank, probably means the user pressed the delete key on a large selection
    {
      const customerQuantityDataSheet = spreadsheet.getSheetByName('Customer Quantity Data');
      var someSKUsNotFound = false, skus;
      var data = customerQuantityDataSheet.getSheetValues(2, 1, customerQuantityDataSheet.getLastRow() - 1, numCols_SearchSheet);

      if (values[0][0].toString().includes(' - ')) // Strip the sku from the first part of the google description
      {
        skus = values.map(item => {
        
          for (var i = 0; i < data.length; i++)
          {
            if (data[i][0].toString().split(' - ').pop().toUpperCase() == item[0].toString().split(' - ').pop().toUpperCase())
              return data[i];
          }

          someSKUsNotFound = true;

          return ['SKU Not Found:', item[0].toString().split(' - ').pop().toUpperCase(), ...new Array(numCols_SearchSheet - 2).fill('')]
        });
      }
      else if (values[0][0].toString().includes('-')) // The SKU contains dashes because that's the convention from Adagio
      {
        skus = values.map(sku => sku[0].substring(0,4) + sku[0].substring(5,9) + sku[0].substring(10)).map(item => {
        
          for (var i = 0; i < data.length; i++)
          {
            if (data[i][0].toString().split(' - ').pop().toUpperCase() == item.toString().toUpperCase())
              return data[i];
          }

          someSKUsNotFound = true;

          return ['SKU Not Found:', item, ...new Array(numCols_SearchSheet - 2).fill('')]
        });
      }
      else // The regular plain SKUs are being pasted
      {
        skus = values.map(item => {
        
          for (var i = 0; i < data.length; i++)
          {
            if (data[i][0].toString().split(' - ').pop().toUpperCase() == item[0].toString().toUpperCase())
              return data[i];
          }

          someSKUsNotFound = true;

          return ['SKU Not Found:', item[0], ...new Array(numCols_SearchSheet - 2).fill('')]
        });
      }

      if (someSKUsNotFound)
      {
        const skusNotFound = [];
        var isSkuFound;

        const skusFound = skus.filter(item => {
          isSkuFound = item[0] !== 'SKU Not Found:'

          if (!isSkuFound)
            skusNotFound.push(item)

          return isSkuFound;
        })

        const numSkusFound = skusFound.length;
        const numSkusNotFound = skusNotFound.length;
        const items = [].concat.apply([], [skusNotFound, skusFound]); // Concatenate all of the item values as a 2-D array
        const numItems = items.length
        const YELLOW = new Array(numCols_SearchSheet).fill('#ffe599');
        const WHITE = new Array(numCols_SearchSheet).fill('white');
        const colours = [].concat.apply([], [new Array(numSkusNotFound).fill(YELLOW), new Array(numSkusFound).fill(WHITE)]); // Concatenate all of the item values as a 2-D array 

        sheet.getRange(5, 1, sheet.getMaxRows() - 4, numCols_SearchSheet).clearContent().setBackground('white').setFontColor('black').setBorder(false, false, false, false, false, false)
          .offset(0, 0, numItems, numCols_SearchSheet).setFontFamily('Arial').setFontWeight('bold').setFontSize(10)
            .setVerticalAlignment('middle').setBackgrounds(colours).setValues(items).setHorizontalAlignment('left').setWrap(true)
          .offset((numSkusFound != 0) ? numSkusNotFound : 0, 0, (numSkusFound != 0) ? numSkusFound : numSkusNotFound, numCols_SearchSheet).activate();

        (numSkusFound !== 1) ? sheet.getRange(1, numCols_SearchSheet).setValue(numSkusFound + " results found.") : sheet.getRange(1, numCols_SearchSheet).setValue(numSkusFound + " result found.");
      }
      else // All SKUs were succefully found
      {
        const numItems = skus.length
        sheet.getRange(5, 1, sheet.getMaxRows() - 4, numCols_SearchSheet).clearContent().setBackground('white').setFontColor('black')
          .offset(0, 0, numItems, numCols_SearchSheet).setFontFamily('Arial').setFontWeight('bold').setFontSize(10)
            .setVerticalAlignment('middle').setBorder(false, false, false, false, false, false).setValues(skus).setHorizontalAlignment('left').setWrap(true).activate();

        (numItems !== 1) ? sheet.getRange(1, numCols_SearchSheet).setValue(numItems + " results found.") : sheet.getRange(1, numCols_SearchSheet).setValue(numItems + " result found.");
      }

      sheet.getRange(2, numCols_SearchSheet).setValue((new Date().getTime() - startTime)/1000 + " s");
    }
  }
}

/**
 * This function searches all of the data for the keywords chosen by the user for the purchase of discovering the invoice numbers that contain the keywords.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf 
 */
function searchForInvoice(e, spreadsheet, sheet)
{
  const startTime = new Date().getTime();
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;
  const rowEnd = range.rowEnd;
  const colEnd = range.columnEnd;

  if (row == 1 && col == 1 && (rowEnd == row || rowEnd == 2) && (col == colEnd || colEnd == null))
  {
    spreadsheet.toast('Searching...', '', -1)
    const YELLOW = "#ffe599";
    const invoiceNumberList = [], highlightedRows = []
    const searchforItems_FilterByCustomer = sheet.getRange(1, 1, 2).clearFormat()                     // Clear the formatting of the range of the search box
      .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
      .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
      .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
      .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
      .getValue().toString().toUpperCase().split(' BY ')                                              // Split the search string at the word 'not'

    const searchforItems_FilterByDate = searchforItems_FilterByCustomer[0].split(' IN ')
    const searchesOrNot = searchforItems_FilterByDate[0].split(' NOT ')
    const searches = searchesOrNot[0].split(' OR ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

    if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
    {
      spreadsheet.toast('Lots of data to parse, this will take about 1 minute.', 'Searching...', -1)

      if (searchforItems_FilterByCustomer.length === 1) // The word 'by' WASN'T found in the string
      {
        if (searchforItems_FilterByDate.length === 1) // The word 'in' wasn't found in the string
        {
          if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
          {
            const numSearches = searches.length; // The number searches
            const dataSheet = spreadsheet.getSheetByName('Invoice Data')
            var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
            var numSearchWords, col_Idx; // Which column of data to search ** Default is item description

            switch (searches[0][0].substring(0, 3)) // Based on the search indicator, set the column of data to search in
            {
              case 'CUS':
                col_Idx = 1;
                searches[0].shift()
                break;
              case 'DAT':
                col_Idx = 2;
                searches[0].shift()
                break;
              case 'INV':
                col_Idx = 3;
                searches[0].shift()
                break;
              case 'LOC':
                col_Idx = 4;
                searches[0].shift()
                break;
              case 'SAL':
                col_Idx = 5;
                searches[0].shift()
                break;
              default:
                col_Idx = 0;
            }

            if (col_Idx < 2 || col_Idx === 3) // Search item descriptions, customer names or invoice numbers
            {
              for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
              {
                loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                {
                  numSearchWords = searches[j].length - 1;

                  for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                  {
                    if (data[i][col_Idx].toString().toUpperCase().includes(searches[j][k])) // Does column 'col_Idx' of the i-th row of data contain the k-th search word in the j-th search
                    {
                      if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                      {
                        if (col_Idx === 0) highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                        if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                        break loop;
                      }
                    }
                    else
                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                  }
                }
              }
            }
            else if (col_Idx === 2) // Search the date
            {
              // The DATE.getMonth() function returns a numeral instead of the name of a month. Use this object to map to the month name.
              const months = {'JAN': 0, 'FEB': 1, 'MAR': 2, 'APR': 3, 'May': 4, 'JUN': 5, 'JUL': 6, 'AUG': 7, 'SEP': 8, 'OCT': 9, 'NOV': 10, 'DEC': 11} 

              for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
              {
                loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                {
                  numSearchWords = searches[j].length;

                  if (numSearchWords === 1) // Assumed to be year only
                  {
                    if (searches[j][0].toString().length === 4) // Check that the year is 4 digits
                    {
                      if (data[i][col_Idx].toString().toUpperCase().includes(searches[j][0])) // Does column 'col_Idx' of the i-th row of data contain the year being searched for
                      {
                        if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                        break loop;
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
                  }
                  else if (numSearchWords === 2) // Assumed to be month and year
                  {
                    if (searches[j][1].toString().length === 4 && searches[j][0].toString().length >= 3) // Check that the year is 4 digits, and the month is atleast 3 characters
                    {     
                      // Does column 'col_Idx' of the i-th row of data contain the year and month being searched for
                      if (data[i][col_Idx].getFullYear() == searches[j][1] && data[i][col_Idx].getMonth() == months[searches[j][0].substring(0, 3)])
                      {
                        if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                        break loop;
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
                  }
                  else if (numSearchWords === 3) // Assumed to be day, month, and year
                  {
                    // Check that the year is 4 digits, the month is atleast 3 characters, and the day is at most 2 characters
                    if (searches[j][2].toString().length === 4 && searches[j][1].toString().length >= 3 && searches[j][0].toString().length <= 2)
                    {
                      // Does column 'col_Idx' of the i-th row of data contain the year, month, and day being searched for
                      if (data[i][col_Idx].getDate() == searches[j][0] && data[i][col_Idx].getFullYear() == searches[j][2] && data[i][col_Idx].getMonth() == months[searches[j][1].substring(0, 3)])
                      {
                        if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                        break loop;
                      }
                      else
                        break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                    }
                  }
                }
              }
            }
            else // Search the location or salesperson ** So much data that we will limit the search to 3 years
            {
              for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
              {
                loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
                {
                  numSearchWords = searches[j].length - 1;

                  for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                  {
                    if (data[i][col_Idx].toString().toUpperCase().includes(searches[j][k])) // Does column 'col_Idx' of the i-th row of data contain the k-th search word in the j-th search
                    {
                      if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                      {
                        if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                        break loop;
                      }
                    }
                    else
                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                  }
                }
              }
            }
          }
          else // The word 'not' was found in the search string
          {
            const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);
            const numWordsToNotInclude = dontIncludeTheseWords.length - 1;
            const numSearches = searches.length; // The number searches
            const dataSheet = spreadsheet.getSheetByName('Invoice Data')
            var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
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
                      for (var l = 0; l <= numWordsToNotInclude; l++) // Loop through the number of words to not include
                      {
                        if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l])) // The i-th description DOES NOT contain the l-th word (of the words that shouldn't be included)
                        {
                          if (l === numWordsToNotInclude) // The i-th description does not include any the words that it is not suppose to
                          {
                            highlightedRows.push(data[i][0]) // Push description
                            if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
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
        }
        else // The word 'in' was found in the string
        {
          const dateSearch = searchforItems_FilterByDate[1].toString().split(" ");

          if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
          {
            const numSearches = searches.length; // The number searches
            const dataSheet = spreadsheet.getSheetByName('Invoice Data')
            var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
            var numSearchWords;

            // The DATE.getMonth() function returns a numeral instead of the name of a month. Use this object to map to the month name.
            const months = {'JAN': 0, 'FEB': 1, 'MAR': 2, 'APR': 3, 'May': 4, 'JUN': 5, 'JUL': 6, 'AUG': 7, 'SEP': 8, 'OCT': 9, 'NOV': 10, 'DEC': 11} 

            for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
            {
              loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
              {
                numSearchWords = searches[j].length - 1;

                for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                {
                  if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does column 1 of the i-th row of data contain the k-th search word in the j-th search
                  {
                    if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                    {
                      if (dateSearch.length === 1) // Assumed to be year only
                      {
                        if (dateSearch[0].toString().length === 4) // Check that the year is 4 digits
                        {
                          if (data[i][2].toString().toUpperCase().includes(dateSearch[0])) // Does column index 2 of the i-th row of data contain the year being searched for
                          {
                            highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                            if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                            break loop;
                          }
                          else
                            break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                        }
                      }
                      else if (dateSearch.length === 2) // Assumed to be month and year
                      {
                        if (dateSearch[1].toString().length === 4 && dateSearch[0].toString().length >= 3) // Check that the year is 4 digits, and the month is atleast 3 characters
                        {     
                          // Does column index of 2 of the i-th row of data contain the year and month being searched for
                          if (data[i][2].getFullYear() == dateSearch[1] && data[i][2].getMonth() == months[dateSearch[0].substring(0, 3)])
                          {
                            highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                            if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                            break loop;
                          }
                          else
                            break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                        }
                      }
                      else if (dateSearch.length === 3) // Assumed to be day, month, and year
                      {
                        // Check that the year is 4 digits, the month is atleast 3 characters, and the day is at most 2 characters
                        if (dateSearch[2].toString().length === 4 && dateSearch[1].toString().length >= 3 && dateSearch[0].toString().length <= 2)
                        {
                          // Does column index of 2 of the i-th row of data contain the year, month, and day being searched for
                          if (data[i][2].getDate() == dateSearch[0] && data[i][2].getFullYear() == dateSearch[2] && data[i][2].getMonth() == months[dateSearch[1].substring(0, 3)])
                          {
                            highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                            if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                            break loop;
                          }
                          else
                            break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                        }
                      }
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
            const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);
            const numWordsToNotInclude = dontIncludeTheseWords.length - 1;
            const numSearches = searches.length; // The number searches
            const dataSheet = spreadsheet.getSheetByName('Invoice Data')
            var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
            var numSearchWords;

            // The DATE.getMonth() function returns a numeral instead of the name of a month. Use this object to map to the month name.
            const months = {'JAN': 0, 'FEB': 1, 'MAR': 2, 'APR': 3, 'May': 4, 'JUN': 5, 'JUL': 6, 'AUG': 7, 'SEP': 8, 'OCT': 9, 'NOV': 10, 'DEC': 11} 

            for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
            {
              loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
              {
                numSearchWords = searches[j].length - 1;

                for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                {
                  if (data[i][0].toString().toUpperCase().includes(searches[j][k])) // Does column 1 of the i-th row of data contain the k-th search word in the j-th search
                  {
                    if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                    {
                      for (var l = 0; l <= numWordsToNotInclude; l++) // Loop through the number of words to not include
                      {
                        if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l])) // The i-th description DOES NOT contain the l-th word (of the words that shouldn't be included)
                        {
                          if (l === numWordsToNotInclude) // The i-th description does not include any the words that it is not suppose to
                          {
                            if (dateSearch.length === 1) // Assumed to be year only
                            {
                              if (dateSearch[0].toString().length === 4) // Check that the year is 4 digits
                              {
                                if (data[i][2].toString().toUpperCase().includes(dateSearch[0])) // Does column index of 2 of the i-th row of data contain the year being searched for
                                {
                                  highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                  if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                  break loop;
                                }
                                else
                                  break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                              }
                            }
                            else if (dateSearch.length === 2) // Assumed to be month and year
                            {
                              if (dateSearch[1].toString().length === 4 && dateSearch[0].toString().length >= 3) // Check that the year is 4 digits, and the month is atleast 3 characters
                              {     
                                // Does column index of 2 of the i-th row of data contain the year and month being searched for
                                if (data[i][2].getFullYear() == dateSearch[1] && data[i][2].getMonth() == months[dateSearch[0].substring(0, 3)])
                                {
                                  highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                  if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                  break loop;
                                }
                                else
                                  break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                              }
                            }
                            else if (dateSearch.length === 3) // Assumed to be day, month, and year
                            {
                              // Check that the year is 4 digits, the month is atleast 3 characters, and the day is at most 2 characters
                              if (dateSearch[2].toString().length === 4 && dateSearch[1].toString().length >= 3 && dateSearch[0].toString().length <= 2)
                              {
                                // Does column index of 2 of the i-th row of data contain the year, month, and day being searched for
                                if (data[i][2].getDate() == dateSearch[0] && data[i][2].getFullYear() == dateSearch[2] && data[i][2].getMonth() == months[dateSearch[1].substring(0, 3)])
                                {
                                  highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                  if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                  break loop;
                                }
                                else
                                  break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                              }
                            }
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
        }
      }
      else // The word 'by' was found in the string
      {
        const customersSearches = searchforItems_FilterByCustomer[1].split(' OR ').map(words => words.split(/\s+/)); // Multiple customers can be searched for

        if (searchforItems_FilterByDate.length === 1) // The word 'in' wasn't found in the string
        {
          if (customersSearches.length === 1) // Search for one customer
          {
            if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
            {
              const numSearches = searches.length; // The number searches
              const numCustomerSearchWords = customersSearches[0].length - 1;
              const dataSheet = spreadsheet.getSheetByName('Invoice Data')
              var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
              var numSearchWords;

              for (var i = 0; i < data.length; i++) // Loop through all of the customers and descriptions from the search data
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
                        for (var l = 0; l <= numCustomerSearchWords; l++) // Loop through each word in the customer search
                        {
                          if (data[i][1].toString().toUpperCase().includes(customersSearches[0][l])) // Does the i-th customer name contain the l-th search word
                          {
                            if (l === numCustomerSearchWords) // All of the customer search words were located in the customer's name
                            {
                              highlightedRows.push(data[i][0]) // Push description
                              if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                              break loop;
                            }
                          }
                          else
                            break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                        }
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
              const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);
              const numWordsToNotInclude = dontIncludeTheseWords.length - 1;
              const numSearches = searches.length; // The number searches
              const numCustomerSearchWords = customersSearches[0].length - 1;
              const dataSheet = spreadsheet.getSheetByName('Invoice Data')
              var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
              var numSearchWords;

              for (var i = 0; i < data.length; i++) // Loop through all of the customers and descriptions from the search data
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
                        for (var l = 0; l <= numWordsToNotInclude; l++) // Loop through the number of words to not include
                        {
                          if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l])) // The i-th description DOES NOT contain the l-th word (of the words that shouldn't be included)
                          {
                            if (l === numWordsToNotInclude) // The i-th description does not include any the words that it is not suppose to
                            {
                              for (var l = 0; l <= numCustomerSearchWords; l++) // Loop through the number of customer search words
                              {
                                if (data[i][1].toString().toUpperCase().includes(customersSearches[0][l])) // The i-th customer name contains the l-th word of the customer search
                                {
                                  if (l === numCustomerSearchWords) // All of the customer search words were located in the customer's name
                                  {
                                    highlightedRows.push(data[i][0]) // Push description
                                    if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                    break loop;
                                  }
                                }
                                else
                                  break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                              }
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
          }
          else // Searching for multiple customers
          {
            if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
            {
              const numSearches = searches.length; // The number searches
              const numCustomerSearches = customersSearches.length; // The number of customer searches
              const dataSheet = spreadsheet.getSheetByName('Invoice Data')
              var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
              var numSearchWords, numCustomerSearchWords;

              for (var i = 0; i < data.length; i++) // Loop through all of the customers and descriptions from the search data
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
                        for (var l = 0; l < numCustomerSearches; l++) // Loop through the number of customer searches
                        {
                          numCustomerSearchWords = customersSearches[l].length - 1;

                          for (var m = 0; m <= numCustomerSearchWords; m++) // Loop through the number of customer search words
                          {
                            if (data[i][1].toString().toUpperCase().includes(customersSearches[l][m])) // Does the i-th customer name contain the m-th search word in the l-th search
                            {
                              if (m === numCustomerSearchWords) // The last customer search word was successfully found in the customer name
                              {
                                highlightedRows.push(data[i][0]) // Push description
                                if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                break loop;
                              }
                            }
                            else
                              break;
                          }
                        }
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
              const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);
              const numWordsToNotInclude = dontIncludeTheseWords.length - 1;
              const numSearches = searches.length; // The number searches
              const numCustomerSearches = customersSearches.length; // The number of customer searches
              const dataSheet = spreadsheet.getSheetByName('Invoice Data')
              var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
              var numSearchWords, numCustomerSearchWords;

              for (var i = 0; i < data.length; i++) // Loop through all of the customers and descriptions from the search data
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
                        for (var l = 0; l <= numWordsToNotInclude; l++) // Loop through the number of words to not include
                        {
                          if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l])) // The i-th description DOES NOT contain the l-th word (of the words that shouldn't be included)
                          {
                            if (l === numWordsToNotInclude) // The i-th description does not include any the words that it is not suppose to
                            {
                              for (var m = 0; m < numCustomerSearches; m++) // Loop through the number of customer searchs
                              {
                                numCustomerSearchWords = customersSearches[m].length - 1;

                                for (var n = 0; n <= numCustomerSearchWords; n++) // Loop through the number of customer search words
                                {
                                  if (data[i][1].toString().toUpperCase().includes(customersSearches[m][n])) // Does the i-th customer name contain the n-th search word in the m-th search
                                  {
                                    if (n === numCustomerSearchWords) // The last customer search word was successfully found in the customer name
                                    {
                                      highlightedRows.push(data[i][0]) // Push description
                                      if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                      break loop;
                                    }
                                  }
                                  else
                                    break;
                                }
                              }
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
          }
        }
        else // The word 'in' was found in the string
        {
          const dateSearch = searchforItems_FilterByDate[1].toString().split(" ");

          if (customersSearches.length === 1) // Search for one customer
          {
            if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
            {
              const numSearches = searches.length; // The number searches
              const numCustomerSearchWords = customersSearches[0].length - 1;
              const dataSheet = spreadsheet.getSheetByName('Invoice Data')
              var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
              var numSearchWords;

              // The DATE.getMonth() function returns a numeral instead of the name of a month. Use this object to map to the month name.
              const months = {'JAN': 0, 'FEB': 1, 'MAR': 2, 'APR': 3, 'May': 4, 'JUN': 5, 'JUL': 6, 'AUG': 7, 'SEP': 8, 'OCT': 9, 'NOV': 10, 'DEC': 11} 

              for (var i = 0; i < data.length; i++) // Loop through all of the customers and descriptions from the search data
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
                        for (var l = 0; l <= numCustomerSearchWords; l++) // Loop through each word in the customer search
                        {
                          if (data[i][1].toString().toUpperCase().includes(customersSearches[0][l])) // Does the i-th customer name contain the l-th search word
                          {
                            if (l === numCustomerSearchWords) // All of the customer search words were located in the customer's name
                            {
                              if (dateSearch.length === 1) // Assumed to be year only
                              {
                                if (dateSearch[0].toString().length === 4) // Check that the year is 4 digits
                                {
                                  if (data[i][2].toString().toUpperCase().includes(dateSearch[0])) // Does column index of 2 of the i-th row of data contain the year being searched for
                                  {
                                    highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                    if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                    break loop;
                                  }
                                  else
                                    break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                                }
                              }
                              else if (dateSearch.length === 2) // Assumed to be month and year
                              {
                                if (dateSearch[1].toString().length === 4 && dateSearch[0].toString().length >= 3) // Check that the year is 4 digits, and the month is atleast 3 characters
                                {     
                                  // Does column index of 2 of the i-th row of data contain the year and month being searched for
                                  if (data[i][2].getFullYear() == dateSearch[1] && data[i][2].getMonth() == months[dateSearch[0].substring(0, 3)])
                                  {
                                    highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                    if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                    break loop;
                                  }
                                  else
                                    break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                                }
                              }
                              else if (dateSearch.length === 3) // Assumed to be day, month, and year
                              {
                                // Check that the year is 4 digits, the month is atleast 3 characters, and the day is at most 2 characters
                                if (dateSearch[2].toString().length === 4 && dateSearch[1].toString().length >= 3 && dateSearch[0].toString().length <= 2)
                                {
                                  // Does column index of 2 of the i-th row of data contain the year, month, and day being searched for
                                  if (data[i][2].getDate() == dateSearch[0] && data[i][2].getFullYear() == dateSearch[2] && data[i][2].getMonth() == months[dateSearch[1].substring(0, 3)])
                                  {
                                    highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                    if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                    break loop;
                                  }
                                  else
                                    break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                                }
                              }
                            }
                          }
                          else
                            break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                        }
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
              const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);
              const numWordsToNotInclude = dontIncludeTheseWords.length - 1;
              const numSearches = searches.length; // The number searches
              const numCustomerSearchWords = customersSearches[0].length - 1;
              const dataSheet = spreadsheet.getSheetByName('Invoice Data')
              var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
              var numSearchWords;

              // The DATE.getMonth() function returns a numeral instead of the name of a month. Use this object to map to the month name.
              const months = {'JAN': 0, 'FEB': 1, 'MAR': 2, 'APR': 3, 'May': 4, 'JUN': 5, 'JUL': 6, 'AUG': 7, 'SEP': 8, 'OCT': 9, 'NOV': 10, 'DEC': 11} 

              for (var i = 0; i < data.length; i++) // Loop through all of the customers and descriptions from the search data
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
                        for (var l = 0; l <= numWordsToNotInclude; l++) // Loop through the number of words to not include
                        {
                          if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l])) // The i-th description DOES NOT contain the l-th word (of the words that shouldn't be included)
                          {
                            if (l === numWordsToNotInclude) // The i-th description does not include any the words that it is not suppose to
                            {
                              for (var l = 0; l <= numCustomerSearchWords; l++) // Loop through each word in the customer search
                              {
                                if (data[i][1].toString().toUpperCase().includes(customersSearches[0][l])) // Does the i-th customer name contain the l-th search word
                                {
                                  if (l === numCustomerSearchWords) // All of the customer search words were located in the customer's name
                                  {
                                    if (dateSearch.length === 1) // Assumed to be year only
                                    {
                                      if (dateSearch[0].toString().length === 4) // Check that the year is 4 digits
                                      {
                                        if (data[i][2].toString().toUpperCase().includes(dateSearch[0])) // Does column index of 2 of the i-th row of data contain the year being searched for
                                        {
                                          highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                          if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                          break loop;
                                        }
                                        else
                                          break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                                      }
                                    }
                                    else if (dateSearch.length === 2) // Assumed to be month and year
                                    {
                                      if (dateSearch[1].toString().length === 4 && dateSearch[0].toString().length >= 3) // Check that the year is 4 digits, and the month is atleast 3 characters
                                      {     
                                        // Does column index of 2 of the i-th row of data contain the year and month being searched for
                                        if (data[i][2].getFullYear() == dateSearch[1] && data[i][2].getMonth() == months[dateSearch[0].substring(0, 3)])
                                        {
                                          highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                          if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                          break loop;
                                        }
                                        else
                                          break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                                      }
                                    }
                                    else if (dateSearch.length === 3) // Assumed to be day, month, and year
                                    {
                                      // Check that the year is 4 digits, the month is atleast 3 characters, and the day is at most 2 characters
                                      if (dateSearch[2].toString().length === 4 && dateSearch[1].toString().length >= 3 && dateSearch[0].toString().length <= 2)
                                      {
                                        // Does column index of 2 of the i-th row of data contain the year, month, and day being searched for
                                        if (data[i][2].getDate() == dateSearch[0] && data[i][2].getFullYear() == dateSearch[2] && data[i][2].getMonth() == months[dateSearch[1].substring(0, 3)])
                                        {
                                          highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                          if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                          break loop;
                                        }
                                        else
                                          break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                                      }
                                    }
                                  }
                                }
                                else
                                  break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                              }
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
          }
          else // Searching for multiple customers
          {
            if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
            {
              const numSearches = searches.length; // The number searches
              const numCustomerSearches = customersSearches.length; // The number of customer searches
              const dataSheet = spreadsheet.getSheetByName('Invoice Data')
              var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
              var numSearchWords, numCustomerSearchWords;

              // The DATE.getMonth() function returns a numeral instead of the name of a month. Use this object to map to the month name.
              const months = {'JAN': 0, 'FEB': 1, 'MAR': 2, 'APR': 3, 'May': 4, 'JUN': 5, 'JUL': 6, 'AUG': 7, 'SEP': 8, 'OCT': 9, 'NOV': 10, 'DEC': 11} 

              for (var i = 0; i < data.length; i++) // Loop through all of the customers and descriptions from the search data
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
                        for (var l = 0; l < numCustomerSearches; l++) // Loop through the number of customer searches
                        {
                          numCustomerSearchWords = customersSearches[l].length - 1;

                          for (var m = 0; m <= numCustomerSearchWords; m++) // Loop through the number of customer search words
                          {
                            if (data[i][1].toString().toUpperCase().includes(customersSearches[l][m])) // Does the i-th customer name contain the m-th search word in the l-th search
                            {
                              if (m === numCustomerSearchWords) // The last customer search word was successfully found in the customer name
                              {
                                if (dateSearch.length === 1) // Assumed to be year only
                                {
                                  if (dateSearch[0].toString().length === 4) // Check that the year is 4 digits
                                  {
                                    if (data[i][2].toString().toUpperCase().includes(dateSearch[0])) // Does column index of 2 of the i-th row of data contain the year being searched for
                                    {
                                      highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                      if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                      break loop;
                                    }
                                    else
                                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                                  }
                                }
                                else if (dateSearch.length === 2) // Assumed to be month and year
                                {
                                  if (dateSearch[1].toString().length === 4 && dateSearch[0].toString().length >= 3) // Check that the year is 4 digits, and the month is atleast 3 characters
                                  {     
                                    // Does column index of 2 of the i-th row of data contain the year and month being searched for
                                    if (data[i][2].getFullYear() == dateSearch[1] && data[i][2].getMonth() == months[dateSearch[0].substring(0, 3)])
                                    {
                                      highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                      if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                      break loop;
                                    }
                                    else
                                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                                  }
                                }
                                else if (dateSearch.length === 3) // Assumed to be day, month, and year
                                {
                                  // Check that the year is 4 digits, the month is atleast 3 characters, and the day is at most 2 characters
                                  if (dateSearch[2].toString().length === 4 && dateSearch[1].toString().length >= 3 && dateSearch[0].toString().length <= 2)
                                  {
                                    // Does column index of 2 of the i-th row of data contain the year, month, and day being searched for
                                    if (data[i][2].getDate() == dateSearch[0] && data[i][2].getFullYear() == dateSearch[2] && data[i][2].getMonth() == months[dateSearch[1].substring(0, 3)])
                                    {
                                      highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                      if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                      break loop;
                                    }
                                    else
                                      break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                                  }
                                }
                              }
                            }
                            else
                              break;
                          }
                        }
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
              const dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);
              const numWordsToNotInclude = dontIncludeTheseWords.length - 1;
              const numSearches = searches.length; // The number searches
              const numCustomerSearches = customersSearches.length; // The number of customer searches
              const dataSheet = spreadsheet.getSheetByName('Invoice Data')
              var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 8);
              var numSearchWords, numCustomerSearchWords;

              // The DATE.getMonth() function returns a numeral instead of the name of a month. Use this object to map to the month name.
              const months = {'JAN': 0, 'FEB': 1, 'MAR': 2, 'APR': 3, 'May': 4, 'JUN': 5, 'JUL': 6, 'AUG': 7, 'SEP': 8, 'OCT': 9, 'NOV': 10, 'DEC': 11} 

              for (var i = 0; i < data.length; i++) // Loop through all of the customers and descriptions from the search data
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
                        for (var l = 0; l <= numWordsToNotInclude; l++) // Loop through the number of words to not include
                        {
                          if (!data[i][0].toString().toUpperCase().includes(dontIncludeTheseWords[l])) // The i-th description DOES NOT contain the l-th word (of the words that shouldn't be included)
                          {
                            if (l === numWordsToNotInclude) // The i-th description does not include any the words that it is not suppose to
                            {
                              for (var m = 0; m < numCustomerSearches; m++) // Loop through the number of customer searchs
                              {
                                numCustomerSearchWords = customersSearches[m].length - 1;

                                for (var n = 0; n <= numCustomerSearchWords; n++) // Loop through the number of customer search words
                                {
                                  if (data[i][1].toString().toUpperCase().includes(customersSearches[m][n])) // Does the i-th customer name contain the n-th search word in the m-th search
                                  {
                                    if (n === numCustomerSearchWords) // The last customer search word was successfully found in the customer name
                                    {
                                      if (dateSearch.length === 1) // Assumed to be year only
                                      {
                                        if (dateSearch[0].toString().length === 4) // Check that the year is 4 digits
                                        {
                                          if (data[i][2].toString().toUpperCase().includes(dateSearch[0])) // Does column index of 2 of the i-th row of data contain the year being searched for
                                          {
                                            highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                            if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                            break loop;
                                          }
                                          else
                                            break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                                        }
                                      }
                                      else if (dateSearch.length === 2) // Assumed to be month and year
                                      {
                                        if (dateSearch[1].toString().length === 4 && dateSearch[0].toString().length >= 3) // Check that the year is 4 digits, and the month is atleast 3 characters
                                        {     
                                          // Does column index of 2 of the i-th row of data contain the year and month being searched for
                                          if (data[i][2].getFullYear() == dateSearch[1] && data[i][2].getMonth() == months[dateSearch[0].substring(0, 3)])
                                          {
                                            highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                            if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                            break loop;
                                          }
                                          else
                                            break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                                        }
                                      }
                                      else if (dateSearch.length === 3) // Assumed to be day, month, and year
                                      {
                                        // Check that the year is 4 digits, the month is atleast 3 characters, and the day is at most 2 characters
                                        if (dateSearch[2].toString().length === 4 && dateSearch[1].toString().length >= 3 && dateSearch[0].toString().length <= 2)
                                        {
                                          // Does column index of 2 of the i-th row of data contain the year, month, and day being searched for
                                          if (data[i][2].getDate() == dateSearch[0] && data[i][2].getFullYear() == dateSearch[2] && data[i][2].getMonth() == months[dateSearch[1].substring(0, 3)])
                                          {
                                            highlightedRows.push(data[i][0]) // Push description if we are doing a regular item search
                                            if (!invoiceNumberList.includes(data[i][3])) invoiceNumberList.push(data[i][3]); // Add the invoice number to the list (if it is not already there)
                                            break loop;
                                          }
                                          else
                                            break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                                        }
                                      }
                                    }
                                  }
                                  else
                                    break;
                                }
                              }
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
          }
        }
      }

      if (invoiceNumberList.length !== 0)
      {
        var output = data.filter(value => invoiceNumberList.includes(value[3]))
        var numItems = output.length;
        var numFormats = new Array(numItems).fill(['@', '@', 'dd MMM yyyy', '@', '@', '@', '@', '$#,##0.00'])

        var backgrounds = output.map(description => {
          if (highlightedRows.includes(description[0]))
            return [YELLOW, YELLOW, YELLOW, YELLOW, YELLOW, YELLOW, YELLOW, YELLOW]
          else
            return ['White', 'White', 'White', 'White', 'White', 'White', 'White', 'White']
        })
      }
      else
      {
        var output = [];
        var numItems = 0;
      }

      if (numItems === 0) // No items were found
      {
        sheet.getRange('A1').activate(); // Move the user back to the seachbox
        sheet.getRange(4, 1, sheet.getMaxRows() - 3, 8).clearContent().setBackground('white'); // Clear content
        const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
        const message = SpreadsheetApp.newRichTextValue().setText("No results found.\nPlease try again.").setTextStyle(0, 16, textStyle).build();
        sheet.getRange(1, 6).setRichTextValue(message);
      }
      else
      {
        sheet.getRange('A4').activate(); // Move the user to the top of the search items
        sheet.getRange(4, 1, sheet.getMaxRows() - 3, 8).clearContent().setBackground('white'); // Clear content and reset the text format
        sheet.getRange(4, 1, numItems, 8).setNumberFormats(numFormats).setBackgrounds(backgrounds).setValues(output);
        const numHighlightedRows = highlightedRows.length;

        if (numItems !== 1)
        {
          if (numHighlightedRows > 1)
            sheet.getRange(1, 6).setValue(numHighlightedRows + ' results found.\n\n' + numItems + ' total rows.')
          else if (numHighlightedRows === 1)
            sheet.getRange(1, 6).setValue('1 result found.\n\n' + numItems + ' total rows.')
          else
            sheet.getRange(1, 6).setValue(numItems +  ' results found.')
        }
        else
          sheet.getRange(1, 6).setValue('1 result found.')
      }

      spreadsheet.toast('Searching Complete.')
    }
    else
    {
      sheet.getRange(4, 1, sheet.getMaxRows() - 3, 8).clearContent(); // Clear content 
      const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
      const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\nPlease try again.").setTextStyle(0, 14, textStyle).build();
      sheet.getRange(1, 6).setRichTextValue(message);
    }

    sheet.getRange(2, 6).setValue((new Date().getTime() - startTime)/1000 + " seconds");
  }
}

/**
 * This function searches for either the amount or quantity of product sold to a particular set of customers, 
 * based on which option the user has selected from the checkboxes on the search sheet.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf 
 */
function searchForQuantityOrAmount(e, spreadsheet, sheet)
{
  const startTime = new Date().getTime();
  const range = e.range;
  const row = range.rowStart;
  const rowEnd = range.rowEnd;

  if (row == rowEnd) // Check and make sure only a single row is being edited //  (rowEnd == row || ***rowEnd == 2***)??
  {
    const col = range.columnStart;
    const colEnd = range.columnEnd;

    if (col == colEnd || colEnd == null) // Check and make sure only a single column is being edited 
    {
      const numYears = new Date().getFullYear() - 2012 + 1;
      var doSearch = false;

      if (row == 1 && col == 1) // The search box was editted
        doSearch = true;
      else 
      {
        if (col == numYears && (row == 2 || row == 3)) // One of the checkboxes were selected
        {
          sheet.getRange(5 - row, numYears).uncheck()
          doSearch = true;
        }
      }

      if (doSearch)
      {
        spreadsheet.toast('Searching...', '', -1)
        const numCols_SearchSheet = sheet.getLastColumn()
        const checkboxes = sheet.getSheetValues(2, numYears, 2, 1);
        var output = [];
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
          if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
          {
            if (searches[0][0] === 'TOP' && /^\d+$/.test(searches[0][1]))
            {
              const dataSheet = selectDataSheet(spreadsheet, checkboxes);

              if (searches[0].length !== 2 && searches[0][2] === 'IN' && /^\d+$/.test(searches[0][3]))
              {
                const y = new Date().getFullYear() + 4 - Number(searches[0][3]);
                output = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).sort((a, b) => b[y] - a[y]);
              }
              else
                output = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).sort((a, b) => b[4] - a[4]);

              output.splice(searches[0][1]);
            }
            else
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
            sheet.getRange(6, 1, sheet.getMaxRows() - 5, numCols_SearchSheet).clearContent(); // Clear content
            const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
            const message = SpreadsheetApp.newRichTextValue().setText("No results found.\nPlease try again.").setTextStyle(0, 16, textStyle).build();
            sheet.getRange(1, numCols_SearchSheet - 3).setRichTextValue(message);
          }
          else
          {
            const horizontalAlignments = new Array(numItems).fill(['left', ...new Array(numCols_SearchSheet - 1).fill('center')]);
            const numFormats = (checkboxes[1][0]) ? new Array(numItems).fill(['@', '@', ...new Array(numCols_SearchSheet - 2).fill('$#,##0.00')]) : new Array(numItems).fill([...new Array(numCols_SearchSheet).fill('@')]);
            sheet.getRange('A6').activate(); // Move the user to the top of the search items
            sheet.getRange(6, 1, sheet.getMaxRows() - 5, numCols_SearchSheet).clearContent().setBackground('white'); // Clear content and reset the text format
            sheet.getRange(6, 1, numItems, numCols_SearchSheet).setNumberFormats(numFormats).setHorizontalAlignments(horizontalAlignments).setValues(output);
            (numItems !== 1) ? sheet.getRange(1, numCols_SearchSheet - 3).setValue(numItems + " results found.") : sheet.getRange(1, numCols_SearchSheet - 3).setValue("1 result found.");
          }

          spreadsheet.toast('Searching Complete.')
        }
        else
        {
          sheet.getRange(6, 1, sheet.getMaxRows() - 5, numCols_SearchSheet).clearContent(); // Clear content 
          const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
          const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\nPlease try again.").setTextStyle(0, 14, textStyle).build();
          sheet.getRange(1, numCols_SearchSheet - 3).setRichTextValue(message);
        }

        sheet.getRange(2, numCols_SearchSheet - 3).setValue((new Date().getTime() - startTime)/1000 + " seconds");
      }
    }
  }
  else if (row > 5)
  {
    spreadsheet.toast('Searching...', '', -1)
    const numCols_SearchSheet = sheet.getLastColumn()
    sheet.getRange(1, 1, 3).clearContent() // Clear the content for the range of the search box
    const values = range.getValues().filter(blank => isNotBlank(blank[0]))

    if (values.length !== 0) // Don't run function if every value is blank, probably means the user pressed the delete key on a large selection
    {
      const numYears = new Date().getFullYear() - 2012 + 1;
      const checkboxes = sheet.getSheetValues(2, numYears, 2, 1);
      const dataSheet = selectDataSheet(spreadsheet, checkboxes);
      var someSKUsNotFound = false, skus;
      var data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, numCols_SearchSheet);

      if (values[0][0].toString().includes(' - ')) // Strip the sku from the first part of the google description
      {
        skus = values.map(item => {
        
          for (var i = 0; i < data.length; i++)
          {
            if (data[i][0].toString().split(' - ').pop().toUpperCase() == item[0].toString().split(' - ').pop().toUpperCase())
              return data[i];
          }

          someSKUsNotFound = true;

          return ['SKU Not Found:', item[0].toString().split(' - ').pop().toUpperCase(), ...new Array(numCols_SearchSheet - 2).fill('')]
        });
      }
      else if (values[0][0].toString().includes('-')) // The SKU contains dashes because that's the convention from Adagio
      {
        skus = values.map(sku => sku[0].substring(0,4) + sku[0].substring(5,9) + sku[0].substring(10)).map(item => {
        
          for (var i = 0; i < data.length; i++)
          {
            if (data[i][0].toString().split(' - ').pop().toUpperCase() == item.toString().toUpperCase())
              return data[i];
          }

          someSKUsNotFound = true;

          return ['SKU Not Found:', item, ...new Array(numCols_SearchSheet - 2).fill('')]
        });
      }
      else // The regular plain SKUs are being pasted
      {
        skus = values.map(item => {
        
          for (var i = 0; i < data.length; i++)
          {
            if (data[i][0].toString().split(' - ').pop().toUpperCase() == item[0].toString().toUpperCase())
              return data[i];
          }

          someSKUsNotFound = true;

          return ['SKU Not Found:', item[0], ...new Array(numCols_SearchSheet - 2).fill('')]
        });
      }

      if (someSKUsNotFound)
      {
        const skusNotFound = [];
        var isSkuFound;

        const skusFound = skus.filter(item => {
          isSkuFound = item[0] !== 'SKU Not Found:'

          if (!isSkuFound)
            skusNotFound.push(item)

          return isSkuFound;
        })

        const numSkusFound = skusFound.length;
        const numSkusNotFound = skusNotFound.length;
        const items = [].concat.apply([], [skusNotFound, skusFound]); // Concatenate all of the item values as a 2-D array
        const numItems = items.length
        const fontColours = new Array(numItems).fill(['black', 'black', 'black', 'black', ...new Array(numCols_SearchSheet - 4).fill('#666666')]);
        const strategies = new Array(numItems).fill([SpreadsheetApp.WrapStrategy.WRAP, ...new Array(numCols_SearchSheet - 1).fill(SpreadsheetApp.WrapStrategy.OVERFLOW)]);
        const YELLOW = new Array(numCols_SearchSheet).fill('#ffe599');
        const WHITE = new Array(numCols_SearchSheet).fill('white');
        const colours = [].concat.apply([], [new Array(numSkusNotFound).fill(YELLOW), new Array(numSkusFound).fill(WHITE)]); // Concatenate all of the item values as a 2-D array
        const horizontalAlignments = [].concat.apply([], [new Array(numSkusNotFound).fill(['left', 'left', ...new Array(numCols_SearchSheet - 2).fill('center')]), 
                                                          new Array(numSkusFound).fill(['left', ...new Array(numCols_SearchSheet - 1).fill('center')])]); 

        sheet.getRange(6, 1, sheet.getMaxRows() - 5, numCols_SearchSheet).clearContent().setBackground('white').setBorder(false, false, false, false, false, false)
          .offset(0, 0, numItems, numCols_SearchSheet).setFontColors(fontColours).setFontFamily('Arial').setFontWeight('bold').setFontSize(10)
            .setHorizontalAlignments(horizontalAlignments).setVerticalAlignment('middle').setBackgrounds(colours).setValues(items).setWrapStrategies(strategies)
          .offset((numSkusFound != 0) ? numSkusNotFound : 0, 0, (numSkusFound != 0) ? numSkusFound : numSkusNotFound, numCols_SearchSheet).activate();

        (numSkusFound !== 1) ? sheet.getRange(1, numCols_SearchSheet - 3).setValue(numSkusFound + " results found.") : sheet.getRange(1, numCols_SearchSheet - 3).setValue(numSkusFound + " result found.");
      }
      else // All SKUs were succefully found
      {
        const numItems = skus.length
        const fontColours = new Array(numItems).fill(['black', 'black', 'black', 'black', ...new Array(numCols_SearchSheet - 4).fill('#666666')]);
        const horizontalAlignments = new Array(numItems).fill(['left', ...new Array(numCols_SearchSheet - 1).fill('center')]);
        const strategies = new Array(numItems).fill([SpreadsheetApp.WrapStrategy.WRAP, ...new Array(numCols_SearchSheet - 1).fill(SpreadsheetApp.WrapStrategy.OVERFLOW)]);
        sheet.getRange(6, 1, sheet.getMaxRows() - 5, numCols_SearchSheet).clearContent().setBackground('white')
          .offset(0, 0, numItems, numCols_SearchSheet).setFontColors(fontColours).setFontFamily('Arial').setFontWeight('bold').setFontSize(10).setHorizontalAlignments(horizontalAlignments)
            .setVerticalAlignment('middle').setBorder(false, false, false, false, false, false).setValues(skus).setWrapStrategies(strategies).activate();

        (numItems !== 1) ? sheet.getRange(1, numCols_SearchSheet - 3).setValue(numItems + " results found.") : sheet.getRange(1, numCols_SearchSheet - 3).setValue(numItems + " result found.");
      }

      sheet.getRange(2, numCols_SearchSheet - 3).setValue((new Date().getTime() - startTime)/1000 + " s");
    }
  }
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
 * This function sorts the invoice data of a particular year via the date first, then sub-sorts lines with the same date via their invoice number.
 */
function sortByDateThenInvoiceNumber(a, b)
{
  return (a[2] > b[2]) ? 1 : (a[2] < b[2]) ? -1 : (a[3] > b[3]) ? 1 : (a[3] < b[3]) ? -1 : 0;
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
        var indexAdjustment = 1999

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