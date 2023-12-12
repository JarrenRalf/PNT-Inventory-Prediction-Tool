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
 * @param {Event Object} e : The event object from an installed onChange trigger.
 */
function installedOnEdit(e)
{
  const range = e.range;
  const col = range.columnStart;
  const row = range.rowStart;
  const rowEnd = range.rowEnd;
  const isSingleRow = row == rowEnd;
  const isSingleColumn = col == range.columnEnd;
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();
  const sheetName = sheet.getSheetName();

  if (sheetName === 'Search for Item Quantity or Amount ($)')
  {
    conditional: if (isSingleColumn)
    {
      if (row == 1 && col == 1 && (rowEnd == 16 || rowEnd == 1))
        searchForQuantityOrAmount(spreadsheet, sheet)
      else if (isSingleRow)
      {
        if (row == 2 && col == 12) // Quantity Data 
          sheet.getRange(3, 12).uncheck()
        else if (row == 3 && col == 12) // Amount ($) Data
          sheet.getRange(2, 12).uncheck()
        else
          break conditional;

        searchForQuantityOrAmount(spreadsheet, sheet)
      }
    }
  }
}

/**
 * This function ... NEEDS A RE_WRITE.
 * 
 * @author Jarren Ralf
 */
function concatenateAllItemData()
{
  const spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.toast('This may take several minutes...', 'Beginning Data Collection')
  const currentYear = new Date().getFullYear();
  const numYears = currentYear - 2012 + 1;
  const years = new Array(numYears).fill('').map((_, y) => (currentYear - y).toString()).reverse(); // Years in ascending order
  const COL = numYears + 4; // A column index to ensure the correct year is being updated when mapping through each year
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
  var quanityData = [], amountData = [], sheet, index, year_y;

  // Loop through all of the years
  years.map((year, y) => {
    spreadsheet.toast('', year)
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

  var N; // The number of terms in the average
  var n; // The index position used to determine the number of terms in the average
  var totalInventory;

  quanityData = quanityData.map((item, i) => {
    n = 1;

    while (isQtyZero(item[item.length - n]))
      n++;

    N = numYears - n + 1;

    if (N > 1) // Compute thew average and make a prediction if we have more than 1 year of data
    {
      [item[3], amountData[i][3]] = getTwoPredictionsUsingLinearRegresssion(
        years.filter((_, y) => y + 1 >= n), // xData
        item.filter((_, t) => t > 4 && t - 5 < N).reverse(),  // yData1
        amountData[i].filter((_, a) => a > 4 && a - 5 < N).reverse(), //yData2
        2024 // X value to predict
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

    adagioInfo = csvData.find(sku => item[0] == sku[6]);
    totalInventory = Number(adagioInfo[2]) + Number(adagioInfo[3]) + Number(adagioInfo[4]) + Number(adagioInfo[5]); // adagioInfo[4] is Trites (400) location; Should we add it in??

    item[1] = adagioInfo[1]; // Update the Adagio description
    item[2] = totalInventory;
    item[4] = Math.round(item[4]*10/N)/10; // Average
    item = item.map(qty => (isQtyNotZero(qty)) ? qty : '') // Remove the zeros, '0', and replace them with a blank string (makes the data present cleaner)

    amountData[i][2] = totalInventory; // Current Inventory
    amountData[i][4] = Math.round(amountData[i][4]*100/N)/100; // Average
    amountData[i] = amountData[i].map(qty => (isQtyNotZero(qty)) ? qty : '') // Remove the zeros

    return item
  })

  spreadsheet.toast('', 'Computations complete...', 120)

  const header = ['SKU', 'Descriptions', 'Current Inventory', 'Prediction', 'Average', ...years.reverse()];
  const numRows_AllQty = quanityData.unshift(header)
  const numRows_AllAmt = amountData.unshift(header)
  const quantityDataSheet = spreadsheet.getSheetByName('Quantity Data');
  const   amountDataSheet = spreadsheet.getSheetByName('Amount Data');

  quantityDataSheet.clear().getRange(1, 1, numRows_AllQty, quanityData[0].length).setValues(quanityData);
  amountDataSheet  .clear().getRange(1, 1, numRows_AllAmt,  amountData[0].length).setValues(amountData);
  quantityDataSheet.deleteColumn(1); // Delete SKU column
    amountDataSheet.deleteColumn(1); // Delete SKU column
  spreadsheet.getSheetByName('Search for Item Quantity or Amount ($)').getRange(1, 10).setValue('Data was last updated on:\n\n' + new Date().toDateString()).offset(0, -9).activate();
  spreadsheet.toast('All Amount / Quantity data has been updated.', 'COMPLETE', 60)
}

/**
 * This function displays the instructions for updating the Search for Item Quantity or Amount ($) data.
 * 
 * @author Jarren Ralf
 */
function display_QuantityOrAmount_Instructions()
{
  showSidebar('Instructions_QuantityOrAmount', 'Update Search for Item Quantity or Amount');
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
    var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3, isYearlyItemData = 4;

    for (var sheet = sheets.length - 1; sheet >= 0; sheet--) // Loop through all of the sheets in this spreadsheet and find the new one
    {
      if (sheets[sheet].getType() == SpreadsheetApp.SheetType.GRID) // Some sheets in this spreadsheet are OBJECT sheets because they contain full charts
      {
        info = [
          sheets[sheet].getLastRow(),
          sheets[sheet].getLastColumn(),
          sheets[sheet].getMaxRows(),
          sheets[sheet].getMaxColumns(),
          sheets[sheet].getRange(1, 3).getValue().toString().includes('Quantity Specif')  // A characteristic of the customer item data
        ]
      
        // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
        if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || 
            (info[maxRow] === 1000 && info[maxCol] === 26 && info[numRows] !== 0 && info[numCols] !== 0) ||
            info[isYearlyItemData]) 
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
            const year = updateYearlyItemData(values, fileName, doesPreviousSheetExist, spreadsheet)
            // concatenateAllItemData()
            // updateAllItemData(year)
            // spreadsheet.getSheetByName('Search for Item Quantity or Amount ($)').getRange(1, 1).activate()
            // spreadsheet.toast('The data will be updated in less than 5 minutes.', 'Import Complete.')
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
 * This function receives the yearly invoice data and it removes the non-imformative SKU numbers, such as the fishing tackle, freight, and marine sale SKUs.
 * 
 * @param {String[][]} data : The yearly invoice data.
 * @return {String[][]} The yearly invoice data with non-imformative SKUs filtered out.
 * @author Jarren Ralf
 */
function removeNonImformativeSKUs(data)
{
  const fishingTackleSKUs = ["80000129", "80000389", "80000549", "80000349", "80000399", "80000499", "80000799", "80000409", "80000439", "80000599", "80000199", "80000249", "80000459", "80000699", "80000739", "80000999", "80001099", "80001149", "80001249", "80001499", "80001949", "80001999", "80000039", "80000089", "80000829", "80000259", "80000589", "80000899", "80000299", "80001199", "80001599", "80000649", "80000849", "80000025", "80000169", "80000579", "80000939", "80001299", "80000139", "80000329", "80000519", "80000629", "80000769", "80000015", "80000149", "80001549", "80000049", "80000949", "80001899", "80000020", "80000079", "80000179", "80000989", "80000449", "80000429", "80000099", "80001699", "80001649", "80001799", "80001849", "80000029", "80000339", "80000749", "80001399", "80000189", "80000289", "80000689", "80000069", "80000279", "80000159", "80000859", "80000729", "80000979", "80000059", "80000229", "80000119", "80000209", "80000219", "80000319", "80000359", "80000369", "80000419", "80000529", "80000639", "80000889", "80001749", "80000789", "80000609", "80000509", "80001049", "80000539", "80000659", "80001449", "80000109", "80000489", "80000759", "80000669", "80000469", "80000379", "80000869", "80000479", "80000679", "80000239", "80000719", "80000569", "80000709", "80000309", "80000919", "80001349", "80000879", "80000929", "80000269", "80000819", "80000619", "80000839", "80000959", "7000F6000", "7000F10000", "80002999", "7000F4000", "7000F5000", "7000F7000", "7000F3000", "7000F8000", "7000F20000", "7000F30000", "7000F9000", "80000779", "80000559", '7000M10000', '7000M200000', '7000M100000', '7000M125000', '7000M15000', '7000M150000', '7000M20000', '7000M3000', '7000M30000', '7000M4000', '7000M5000', '7000M50000', '7000M6000', '7000M7000', '7000M75000', '7000M8000', '7000M9000', 'FREIGHT', 'MISCITEM', 'MISCWEB', 'GIFT CERTIFICATE', 'BROKERAGE', 'ROPE SPLICE', '54002800', '54003600', '20110000', '7000C24999', '20120000']

  return data.filter(v => !fishingTackleSKUs.includes(v[9].toString()))
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
function searchForQuantityOrAmount(spreadsheet, sheet)
{
  const startTime = new Date().getTime();
  const searchResultsDisplayRange = sheet.getRange(1, 13); // The range that will display the number of items found by the search
  const functionRunTimeRange = sheet.getRange(2, 13);      // The range that will display the runtimes for the search and formatting
  const itemSearchFullRange = sheet.getRange(6, 1, sheet.getMaxRows() - 5, 16); // The entire range of the Item Search page
  const checkboxes = sheet.getSheetValues(2, 12, 2, 1);
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
      const numCols = dataSheet.getLastColumn()
      const data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, numCols);
      const numSearches = searches.length; // The number searches
      var numSearchWords;

      if (searches[0][0] === 'RECENT') // If the user's search begins with 'RECENT' then they are searching for information in the final column of data, which is the one that contains customer info
      {
        const lastColIndex = numCols - 1;

        if (searches[0].length !== 1) // Also contains a search phrase
        {
          searches[0].shift() // Remove the 'RECENT' keyword
          const searchPhrase = searches[0].join(" ") // Join the rest of the search terms together
          output.push(...data.filter(customer => (isNotBlank(customer[lastColIndex])) ? customer[lastColIndex].includes(searchPhrase) : false)) // If the final column is not blank, find the phrase
        }
        else // Return the last two years of data
          output.push(...data.filter(customer => isNotBlank(customer[lastColIndex])))
      }
      else
      {
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
      const data = dataSheet.getSheetValues(2, 1, dataSheet.getLastRow() - 1, 16);
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
      var numFormats = (checkboxes[1][0]) ? new Array(numItems).fill(['@', '@', ...new Array(14).fill('$#,##0.00')]) : new Array(numItems).fill([...new Array(16).fill('@')]);
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
 * This function displays a side bar.
 * 
 * @param {String} htmlFileName : The name of the html file.
 * @param {String}    title     : The title of the sidebar.
 * @author Jarren Ralf
 */
function showSidebar(htmlFileName, title) 
{
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile(htmlFileName).setTitle(title));
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
 * This function ... NEEDS A RE_WRITE.
 * 
 * @param {Number} year : The year of data that was just imported into the spreadsheet. 
 * @author Jarren Ralf
 */
function updateAllItemData(year)
{
  const spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.toast('This may take several minutes...', 'Beginning Data Collection')

  const currentYear = new Date().getFullYear();
  const numYears = currentYear - 2012 + 1;
  const years = new Array(numYears).fill('').map((_, y) => (currentYear - y).toString()).reverse(); // Years in ascending order
  const COL = numYears + 2; // A column index to ensure the correct year is being updated when mapping through each year
  //const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
  var quanityData = [], amountData = [], sheet, index, item, year_y;

  // Loop through all of the years
  years.map((year, y) => {
    year_y = COL - y; // The appropriate index for the y-th year

    sheet = spreadsheet.getSheetByName(year)
    sheet.getSheetValues(2, 2, sheet.getLastRow() - 1, 5).map(salesData => { // Loop through all of the sales data for the y-th year
      if (isNotBlank(salesData[0])) // Spaces between customers
      {
        index = quanityData.findIndex(d => d[0] === item[1]); // The index for the current item in the combined quantity data

        if (index !== -1) // Current item is already in combined data list
        {
          quanityData[index][year_y] += Number(salesData[3]) // Increase the quantity
            amountData[index][year_y] += Number(salesData[4]) // Increase the amount ($)
        }
        else // The current item is not in the combined data yet, so add it in
        {
          quanityData.push([item[1], 0, 0, ...new Array(numYears).fill(0), ''])
            amountData.push([item[1], 0, 0, ...new Array(numYears).fill(0), ''])
          quanityData[quanityData.length - 1][year_y] = Number(salesData[3]) // Add quantity to the appropriate year (column)
            amountData[amountData.length  - 1][year_y] = Number(salesData[4]) // Add amount ($) to the appropriate year (column)
        }
      }
    })
  })

  quanityData = quanityData.map((item, i) => {
      item[1] = Math.round((item[3] + item[4] + item[5] + item[6] + item[7] + item[8])*5/3)/10; // Average
      item[2] = Math.round((item[3] + item[6] + item[7] + item[8])*5/2)/10; // Average - Covid
      item = item.map(qty => (isQtyNotZero(qty)) ? qty : '') // Remove the zeros, '0', and replace them with a blank string (makes the data present cleaner)
      amountData[i][1] = 
        Math.round((amountData[i][3] + amountData[i][4] + amountData[i][5] + amountData[i][6] + amountData[i][7] + amountData[i][8])*50/3)/100; // Average
      amountData[i][2] =  Math.round((amountData[i][3] + amountData[i][6] + amountData[i][7] + amountData[i][8])*25)/100; // Average - Covid
      amountData[i] = amountData[i].map(qty => (isQtyNotZero(qty)) ? qty : '') // Remove the zeros
      return item
    })

  const header = ['Descriptions', 'AVG (6 yr)', 'AVG - CoV', ...years.reverse(), 'Customers purchased in 2023'];
  const numRows_AllQty = quanityData.unshift(header)
  const numRows_AllAmt = amountData.unshift(header)

  spreadsheet.getSheetByName('Quantity Data').clear().getRange(1, 1, numRows_AllQty, quanityData[0].length).setValues(quanityData)
  spreadsheet.getSheetByName('Amount Data').clear().getRange(1, 1, numRows_AllAmt, amountData[0].length).setValues(amountData)
  spreadsheet.getSheetByName('Search for Item Quantity or Amount ($)').getRange(1, 16, 4)
    .setValues([['Data was last updated on:\n\n' + new Date().toDateString()],[''],[''],
                ['Customers who purchased these items in ' + (currentYear - 1).toString() + ' and ' + currentYear.toString()]])
  spreadsheet.toast('All Amount / Quantity data has been updated.', 'COMPLETE', 60)
}

/**
 * This function ...
 * 
 * @param {Object[][]}      values         : The values of the data that were just imported into the spreadsheet
 * @param {String}         fileName        : The name of the new sheet (which will also happen to be the xlxs file name)
 * @param {Boolean} doesPreviousSheetExist : Whether the previous sheet with the same name exists or not
 * @return {Number}         year           : Returns the year of the data that was just imported into the spreadsheet
 * @author Jarren Ralf
 */
function updateYearlyItemData(values, fileName, doesPreviousSheetExist, spreadsheet)
{
  const currentYear = new Date().getFullYear();
  values.shift() // Remove the header
  values.pop()   // Remove the final row which contains descriptive stats
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
        const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
        const data = values.filter(itemNumber => {
          item = csvData.find(sku => itemNumber[0] == sku[6])

          if (item != undefined && item[10] === 'A') // Item is found and Active
            itemNumber[1] = item[1] // Update the Descriptions with the google descriptons

          return item != undefined && item[10] === 'A'; // Item is found and Active
        })


        const numCols = 4;
        const sheets = spreadsheet.getSheets();
        const previousSheet = sheets.find(sheet => sheet.getSheetName() == year)
        var indexAdjustment = 2012;

        if (doesPreviousSheetExist)
        {
          indexAdjustment--;
          spreadsheet.deleteSheet(previousSheet)
        }
        
        SpreadsheetApp.flush();
        const newSheet = spreadsheet.insertSheet(year, sheets.length + indexAdjustment - year)
          .setColumnWidth(1, 150).setColumnWidth(2, 700).setColumnWidths(3, 2, 75);
        SpreadsheetApp.flush();
        const lastRow = data.unshift(['Item Number', 'Item Description', 'Quantity', 'Amount']);
        newSheet.deleteColumns(5, 22)
        newSheet.setTabColor('#a64d79').setFrozenRows(1)
        newSheet.protect()
        newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').setBackground('#c0c0c0')
          .offset(0, 0, lastRow, numCols).setHorizontalAlignments(new Array(lastRow).fill(['left', 'left', 'right', 'right'])).setNumberFormat('@').setValues(data)
        newSheet.hideSheet();

        return year;
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
    const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString())
    const data = values.filter(itemNumber => {
      item = csvData.find(sku => itemNumber[0] == sku[6])

      if (item != undefined && item[10] === 'A') // Item is found and Active
        itemNumber[1] = item[1] // Update the Descriptions with the google descriptons

      return item != undefined && item[10] === 'A'; // Item is found and Active
    })

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
    const lastRow = data.unshift(['Item Number', 'Item Description', 'Quantity', 'Amount']);
    newSheet.deleteColumns(5, 22)
    newSheet.setTabColor('#a64d79').setFrozenRows(1)
    newSheet.protect()
    newSheet.getRange(1, 1, 1, numCols).setFontSize(11).setFontWeight('bold').setBackground('#c0c0c0')
      .offset(0, 0, lastRow, numCols).setHorizontalAlignments(new Array(lastRow).fill(['left', 'left', 'right', 'right'])).setNumberFormat('@').setValues(data)
    newSheet.hideSheet();

    return year;
  }
}