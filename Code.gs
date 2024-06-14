/**
 * This function handles the on edit events in this spreadsheet pertaining to the Item Search sheet only (all other sheets will be protected).
 * This function is looking for the user searching for items and it is making appropriate changes to the data when a user deletes items from their order.
 * 
 * @param {Event Object} e : The event object
 * @OnlyCurrentDoc
 */
function onEdit(e)
{ 
  const range = e.range;
  const col = range.columnStart;
  const row = range.rowStart;
  const rowEnd = range.rowEnd;
  const isSingleRow = row == rowEnd;
  const isSingleColumn = col == range.columnEnd;
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();

  if (range.getA1Notation() === 'D2')
    Logger.log('*** SUBMISSION BOX EDITTED ***')

  Logger.log('range: ' + range.getA1Notation())
  Logger.log('row: ' + row)
  Logger.log('col: ' + col)
  Logger.log('rowEnd: ' + rowEnd)
  Logger.log('isSingleRow: ' + isSingleRow)
  Logger.log('isSingleColumn: ' + isSingleColumn)

  if (sheet.getSheetName() === 'Item Search' && isSingleColumn)
    if (row == 1 && col == 1 && (rowEnd == null || rowEnd == 2 || isSingleRow))
      search(e, spreadsheet, sheet);
    else if (row == 2 && col == 4) // Submission Checkbox
      checkForOrderSubmission(range, sheet, spreadsheet);
    else if (row > 4) // If the body of the item Search is being edited
      if (col == 6) // Items are being selected in the description column
        deleteItemsFromOrder(sheet, range, range.getValue(), row, isSingleRow, spreadsheet);
      else if (col == 1 || col == 3 || col == 5) // The SKU, UoM, or the Descriptions - Categories - Unit of Measure - SKU # column are being edited (The user is not suppose to edit these fields)
        undoUserMistake(sheet, e, range, isSingleRow, spreadsheet)
}

/**
 * This function identifies all of the cells that the user has selected and moves those items to the order portion of the Item Search sheet.
 * 
 * @author Jarren Ralf
 * @OnlyCurrentDoc
 */
function addSelectedItemsToOrder()
{
  const startTime = new Date().getTime(); // Used for the function runtime
  var firstRows = [], firstCols = [], lastRows = [], lastCols = [], itemValues = [], splitDescription, sku, uom;
  const sheet = SpreadsheetApp.getActiveSheet();

  sheet.getActiveRangeList().getRanges().map((rng, r) => {
    firstRows.push(rng.getRow());
    lastRows.push(rng.getLastRow());
    firstCols.push(rng.getColumn());
    lastCols.push(rng.getLastColumn());
    itemValues.push(...sheet.getSheetValues(firstRows[r], 1, lastRows[r] - firstRows[r] + 1, 1))
  })

  if (Math.min(...firstCols) === Math.max(...lastCols) && Math.min(...firstRows) > 4 && Math.max( ...lastRows) <= sheet.getLastRow()) // If the user has not selected an item, alert them with an error message
  { 
    const numItems = itemValues.length;
    const row = getNumRows(sheet.getSheetValues(5, 3, sheet.getLastRow() - 4, 5)) + 5;
    sheet.getRange(row, 3, numItems, 4).setNumberFormat('@').setValues(itemValues.map(item => {
      splitDescription = item[0].split(' - ');
      sku = splitDescription.pop();
      uom = splitDescription.pop();
      splitDescription.pop();
      return [sku, '', uom, splitDescription]
    })).offset(0, 1, 1, 1).activate(); // Move to the quantity column
  }
  else
    SpreadsheetApp.getUi().alert('Please select an item from the list.');

  sheet.getRange(2, 5).setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * This function retrieves the items on the Recently Created and places them on the Item Search sheet.
 * 
 * @author Jarren Ralf
 * @OnlyCurrentDoc
 */
function allItems()
{
  const startTime = new Date().getTime(); // Used for the function runtime
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = SpreadsheetApp.getActiveSheet();
  const recentlyCreatedSheet = spreadsheet.getSheetByName('Recently Created');
  const numItems = recentlyCreatedSheet.getLastRow();
  sheet.getRange(1, 1).clearContent() // Clear the search box
    .offset(4, 0, sheet.getMaxRows() - 4).clearContent() // Clear the previous search
    .offset(0, 0, numItems).setValues(recentlyCreatedSheet.getSheetValues(1, 1, numItems, 1)) // Set the values
    .offset(-3, 5, 1, 1).setValue("Items displayed in order of newest to oldest.") // Tell user items are sorted from newest to oldest
    .offset(0, -1).setValue((new Date().getTime() - startTime)/1000 + " seconds"); // Function runtime
  spreadsheet.toast('PNT\'s most recently created items are being displayed.');
}

/**
 * This function checks if the customer has checked or unchecked their submission button. In both cases, this will lead to the PNT Order Processor spreadsheet to send the appropriate 
 * emails to the customer and PNT employees.
 * 
 * @param    {Range}       range    : The range of the checkbox.
 * @param    {Sheet}       sheet    : The sheet that was last editted
 * @param {spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 * @OnlyCurrentDoc
 */
function checkForOrderSubmission(range, sheet, spreadsheet)
{
  const startTime = new Date().getTime(); // Used for the function runtime
  Logger.log('checkForOrderSubmission is being run...')
  Logger.log('Is box checked? : ' + range.isChecked())
  range.offset(2, -2).setValue(startTime); // Place the timestamp in one of the hidden cells
  SpreadsheetApp.flush(); // Force the change on the spreadsheet first

  if (range.isChecked())
  {
    const isSubmissionSuccessful = getExportData(sheet, spreadsheet);
    Logger.log('Is submission successful? : ' + isSubmissionSuccessful);
    range.offset(0, 1).setValue((new Date().getTime() - startTime)/1000 + " seconds");
    SpreadsheetApp.flush(); // Force the change on the spreadsheet first
    
    if (isSubmissionSuccessful)
      SpreadsheetApp.getUi().alert('Your order has been submitted.\n\nThank You!');
  }
}

/**
 * This function handles the task of deleting items from the users order on the Item Search sheet. 
 * It finds the missing descriptions and it moves the data up to fill in the gap.
 * 
 * @param {Sheet}          sheet    : The Item Search sheet
 * @param {Range}          range    : The active range
 * @param {String[][]}     value    : The values in the range that were editted
 * @param {Number}          row     : The first row that was editted
 * @param {Boolean}     isSingleRow : Whether or not a single row was editted
 * @param {Spreadsheet} spreadsheet : The active spreadsheet
 * @author Jarren Ralf
 * @OnlyCurrentDoc
 */
function deleteItemsFromOrder(sheet, range, value, row, isSingleRow, spreadsheet)
{
  const startTime = new Date().getTime(); // Used for the function runtime
  spreadsheet.toast('Checking for possible lines to delete...')
  const numRows = getNumRows(sheet.getSheetValues(5, 3, sheet.getLastRow() - 4, 5));

  if (numRows > 0)
  {
    if (isSingleRow)
    {
      if (!Boolean(value)) // Was a single cell editted?, is the value blank? or is the quantity zero?
      {
        const itemsOrderedRange = sheet.getRange(row, 3, numRows - row + 5, 5);
        const orderedItems = itemsOrderedRange.getValues();
        orderedItems.shift(); // This is the item that was deleted by the user
        itemsOrderedRange.clearContent()

        if (orderedItems.length > 0)
          itemsOrderedRange.offset(0, 0, orderedItems.length).setValues(orderedItems); // Move the items up to fill in the gap

        spreadsheet.toast('Deleting Complete.')
      }
      else
        spreadsheet.toast('Nothing Deleted.')
    }
    else if (isEveryValueBlank(range.getValues())) // Multiple rows
    {
      const itemsOrderedRange = sheet.getRange(5, 3, numRows, 5);
      const orderedItems = itemsOrderedRange.getValues().filter(description => isNotBlank(description[3])); // Find and remove the blank descriptions
      itemsOrderedRange.clearContent();
      
      if (orderedItems.length > 0)
        itemsOrderedRange.offset(0, 0, orderedItems.length, 5).setValues(orderedItems); // Move the items up to fill in the gaps 

      spreadsheet.toast('Deleting Complete.')
    }
    else
      spreadsheet.toast('Nothing Deleted.')
  }
  else
    spreadsheet.toast('Nothing Deleted.')

  sheet.getRange(2, 5).setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * This function gets the export data from all of the customer's spreadsheets that have submitted their order.
 * 
 * @param   {Sheet}   itemSearchSheet : The item search sheet
 * @param {Spreadsheet} spreadsheet   : The active spreadsheet
 * @returns Whether or not the export was successful.
 * @throws General error if anything goes wrong
 * @author Jarren Ralf
 */
function getExportData(itemSearchSheet, spreadsheet)
{
  spreadsheet.toast('Your request is being processed...')
  Logger.log('Getting export data...')

  try
  {
    const numRows = getNumRows(itemSearchSheet.getSheetValues(5, 3, itemSearchSheet.getLastRow() - 4, 5));
    
    if (numRows > 0) // The customer's order is not blank
    {
      const range = itemSearchSheet.getRange(5, 3, numRows, 5);
      const values = range.getValues();

      if (!values.every(qty => isBlank(qty[1]))) // Not every order quantity is blank
      {
        const deliveryInstructions = itemSearchSheet.getSheetValues(2, 7, 1, 1)[0][0];
        const poNumber = itemSearchSheet.getSheetValues(1, 6, 1, 1)[0][0];
        const customerAccountNumber = itemSearchSheet.getSheetValues(4, 3, 1, 1)[0][0];
        const submittedOrdersSheet = spreadsheet.getSheetByName('Submitted Orders');
        const lastRow = submittedOrdersSheet.getLastRow();
        const today = new Date().toLocaleString();
        const previousBackgroundColour = submittedOrdersSheet.getRange(lastRow, 1).getBackground();
        const numItems = values.length;
        submittedOrdersSheet.getRange(lastRow + 1, 1, numItems, 5)
          .setBackground((previousBackgroundColour !== '#ffffff') ? 'white' : '#c9daf8')
          .setNumberFormat('@')
          .setHorizontalAlignments(new Array(numItems).fill(['center', 'left', 'right', 'center', 'left']))
          .setValues(values.map(item => [today, poNumber, item[1], item[0], item[3]]));
        spreadsheet.getSheetByName('Last Export').clearContents().getRange(1, 1, numRows + 1, 5).setNumberFormat('@').setValues([['', '', '', poNumber, deliveryInstructions] , ...values]) // Used for email
        SpreadsheetApp.flush();
        const exportData_WithDiscountedPrices = [];

        /* If there are delivery instructions, make them the final line of the order.
         * If necessary, make multiple comment lines if comments are > 75 characters long.
         */
        const exportData = [['H', customerAccountNumber, poNumber, 'PNT DELIVERY'], ...values, // The SKUs and quantities
          ['I', 'Provide your preferred delivery / pick up date and location below:', '', ''],
          ...(isNotBlank(deliveryInstructions)) ? deliveryInstructions.match(/.{1,75}/g).map(c => ['I', c, '', '']) : [['I', '**Customer left this field blank**', '', '']]];

        exportData.map(item => {
          if (item[0] === 'H')
            exportData_WithDiscountedPrices.push(['H', item[1], item[2], item[3]]);
          else if (item[0] === 'I')
            exportData_WithDiscountedPrices.push(['I', item[1], '', '']);
          else // There was no line indicator
          {
            item[0] = item[0].toString().trim().toUpperCase(); // Make the SKU uppercase

            if (isNotBlank(item[0])) // SKU is not blank
              if (isNotBlank(item[1])) // Order quantity is not blank
                if (Number(item[1]).toString() !== 'NaN') // Order number is a valid number
                  exportData_WithDiscountedPrices.push(['D', item[0], 0, item[1]], ...(item[3] + ' - ' + item[2]).toString().match(/.{1,62}/g).map(c => ['C', 'Description: ' + c, '', '']))
                else // Order quantity is not a valid number
                  exportData_WithDiscountedPrices.push(
                    ['D', item[0], 0, 0], 
                    ...(item[3] + ' - ' + item[2]).toString().match(/.{1,62}/g).map(c => ['C', 'Description: ' + c, '', '']),
                    ['C', 'Invalid order QTY: "' + item[1] + '" for above item, therefore it was replaced with 0', '', '']
                  )
              else // The order quantity is blank (while SKU is not)
                exportData_WithDiscountedPrices.push(
                  ['D', item[0], 0, 0],
                  ...(item[3] + ' - ' + item[2]).toString().match(/.{1,62}/g).map(c => ['C', 'Description: ' + c, '', '']),
                  ['C', 'Order quantity was blank for the above item, therefore it was replaced with 0', '', '']
                )
            else // The SKU is blank
              if (isNotBlank(item[1])) // Order quantity is not blank
                if (Number(item[1]).toString() !== 'NaN') // Order number is a valid number
                  exportData_WithDiscountedPrices.push(
                    ['D', 'MISCITEM', 0, item[1]], 
                    ...(item[3] + ' - ' + item[2]).toString().match(/.{1,62}/g).map(c => ['C', 'Description: ' + c, '', ''])
                  )
                else // Order quantity is not a valid number
                  exportData_WithDiscountedPrices.push(
                    ['D', 'MISCITEM', 0, 0], 
                    ...(item[3] + ' - ' + item[2]).toString().match(/.{1,62}/g).map(c => ['C', 'Description: ' + c, '', '']),
                    ['C', 'Invalid order QTY: "' + item[1] + '" for above item, therefore it was replaced with 0', '', '']
                  )
              else // The order quantity is blank 
                if (isNotBlank(item[3])) // Description is not blank (but SKU and quantity are)
                  exportData_WithDiscountedPrices.push(
                    ['D', 'MISCITEM', 0, 0], 
                    ...(item[3] + ' - ' + item[2]).toString().match(/.{1,62}/g).map(c => ['C', 'Description: ' + c, '', '']),
                    ['C', 'Order quantity was blank for the above item, therefore it was replaced with 0', '', '']
                  )

            if (isNotBlank(item[4])) // There are notes for the current line
              exportData_WithDiscountedPrices.push(...(item[4]).match(/.{1,68}/g).map(c => ['C', 'Notes: ' + c, '', '']))
          }
        })

        const exportSheet = spreadsheet.getSheetByName('Export');
        const rowStart = exportSheet.getLastRow() + 1;
        const ranges = [[], [], []];
        const backgroundColours = [
          '#c9daf8', // Make the header rows blue
          '#fcefe1', // Make the comment rows orange
          '#e0d5fd'  // Make the instruction comment rows purple
        ];

        exportData_WithDiscountedPrices.map((h, r) => 
          h = (h[0] !== 'H') ? (h[0] !== 'C') ? (h[0] !== 'I') ? false : 
          ranges[2].push('A' + (r + rowStart) + ':D' + (r + rowStart)) : // Instruction comment rows purple
          ranges[1].push('A' + (r + rowStart) + ':D' + (r + rowStart)) : // Comment rows orange
          ranges[0].push('A' + (r + rowStart) + ':D' + (r + rowStart))   // Header rows blue
        )
        
        exportSheet.getRange(rowStart, 1, exportData_WithDiscountedPrices.length, 4).setNumberFormat('@').setBackground('white').setValues(exportData_WithDiscountedPrices);
        ranges.map((rngs, r) => (rngs.length !== 0) ? exportSheet.getRangeList(rngs).setBackground(backgroundColours[r]) : false); // Set the appropriate background colours
        SpreadsheetApp.flush()

        range.clearContent() // Clear the customers order, including notes
          .offset(-4,  3, 1, 1).setValue('').activate() // Remove the Customer PO #
          .offset( 1,  0, 1, 2).setValues([['Items displayed in order of newest to oldest.', '']]) // Remove the Delivery / Pick Up instructions
          .offset( 0, -2).uncheck() // Uncheck the submit order checkbox

        return true;
      }
      else
      {
        itemSearchSheet.getRange(2, 4).uncheck();
        SpreadsheetApp.flush();
        const ui = SpreadsheetApp.getUi();
        ui.alert('Error: Order Quantities are All Blank', 'Please add an order quantity for each of the items on your order.', ui.ButtonSet.OK);
        return false;
      }
    }
    else
    {
      itemSearchSheet.getRange(2, 4).uncheck();
      SpreadsheetApp.flush();
      const ui = SpreadsheetApp.getUi();
      ui.alert('Error: No Items on Order', 'Please add items your order before submitting.', ui.ButtonSet.OK);
      return false;
    }
  }
  catch (e)
  {
    itemSearchSheet.getRange(2, 4).uncheck();
    SpreadsheetApp.flush();
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error: Submission Failure', 'Please contact Jarren at PNT and let him know that the following error occured during order submission:\n\n' + e['stack'], ui.ButtonSet.OK);
  }
}

/**
 * This function gets the last effective row of the customer order on the Item Search page
 *
 * @param {Object[][]} vals : This is the array of values that will contain the customer's order
 * @returns The number of rows in the order.
 * @author Jarren Ralf
 * @OnlyCurrentDoc
 */ 
function getNumRows(vals)
{
  for (var i = vals.length - 1; i > -1; i--)
    if (isNotBlank(vals[i][0]) || isNotBlank(vals[i][1]) || isNotBlank(vals[i][2]) || isNotBlank(vals[i][3]) || isNotBlank(vals[i][4]))
      return i + 1;
    
  return 0;
}

/**
 * This function checks if the given string is blank or not.
 * 
 * @param {String} str : The given string.
 * @returns {Boolean} Whether the given string is blank or not.
 * @author Jarren Ralf
 * @OnlyCurrentDoc
 */
function isBlank(str)
{
  return str === '';
}

/**
 * This function checks if every value in the import multi-array is blank, which means that the user has
 * highlighted and deleted all of the data.
 * 
 * @param {Object[][]} values : The import data
 * @return {Boolean} Whether the import data is deleted or not
 * @author Jarren Ralf
 * @OnlyCurrentDoc
 */
function isEveryValueBlank(values)
{
  return values.every(arr => arr.every(val => val == '') === true);
}

/**
 * This function checks if the given string is not blank or not.
 * 
 * @param {String} str : The given string.
 * @returns {Boolean} Whether the given string is not blank or not.
 * @author Jarren Ralf
 * @OnlyCurrentDoc
 */
function isNotBlank(str)
{
  return str !== '';
}

/**
 * This function first applies the standard formatting to the search box, then it seaches the Item List page for the items in question.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf 
 * @OnlyCurrentDoc
 */
function search(e, spreadsheet, sheet)
{
  const startTime = new Date().getTime(); // Used for the function runtime
  const output = [];
  const searchesOrNot = sheet.getRange(1, 1, 2).clearFormat()                                       // Clear the formatting of the range of the search box
    .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
    .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
    .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
    .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
    .getValue().toString().toLowerCase().split(' not ')                                             // Split the search string at the word 'not'

  const searches = searchesOrNot[0].split(' or ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

  if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
  {
    spreadsheet.toast('Searching...')

    const inventorySheet = spreadsheet.getSheetByName('Item List');
    const data = inventorySheet.getSheetValues(1, 1, inventorySheet.getLastRow(), 1);
    const numSearches = searches.length; // The number searches
    var numSearchWords;

    if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
    {
      for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
      {
        loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
        {
          numSearchWords = searches[j].length - 1;

          for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
          {
            if (data[i][0].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
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
      var numSearchWords_ToNotInclude;

      for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
      {
        loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
        {
          numSearchWords = searches[j].length - 1;

          for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
          {
            if (data[i][0].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
            {
              if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
              {
                numSearchWords_ToNotInclude = dontIncludeTheseWords.length - 1;

                for (var l = 0; l <= numSearchWords_ToNotInclude; l++)
                {
                  if (!data[i][0].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                  {
                    if (l === numSearchWords_ToNotInclude)
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
      sheet.getRange('A1').activate() // Move the user back to the seachbox
        .offset(4, 0, sheet.getMaxRows() - 4).clearContent() // Clear content
        .offset(-3, 5, 1, 1).setValue("No results found.\nPlease try again.");
    else
      sheet.getRange('A5').activate() // Move the user to the top of the search items
        .offset(0, 0, sheet.getMaxRows() - 4).clearContent()
        .offset(0, 0, numItems).setValues(output) 
        .offset(-3, 5, 1, 1).setValue((numItems !== 1) ? numItems + " results found." : numItems + " result found.");

    spreadsheet.toast('Searching Complete.');
  }
  else if (isNotBlank(e.oldValue) && userHasPressedDelete(e.value)) // If the user deletes the data in the search box, then the recently created items are displayed
  {
    spreadsheet.toast('Accessing most recently created items...');
    const recentlyCreatedItemsSheet = spreadsheet.getSheetByName('Recently Created');
    const numItems = recentlyCreatedItemsSheet.getLastRow();
    sheet.getRange('A5').activate() // Move the user to the top of the search items
      .offset(0, 0, sheet.getMaxRows() - 4).clearContent()
      .offset(0, 0, numItems).setValues(recentlyCreatedItemsSheet.getSheetValues(1, 1, numItems, 1))
      .offset(-3, 5, 1, 1).setValue("Items displayed in order of newest to oldest.")
    spreadsheet.toast('PNT\'s most recently created items are being displayed.')
  }
  else
  {
    sheet.getRange(5, 1, sheet.getMaxRows() - 4).clearContent() // Clear content 
      .offset(-3, 5, 1, 1).setValue("Invalid search.\nPlease try again.");
    spreadsheet.toast('Invalid Search.');
  }

  sheet.getRange(2, 5).setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * This function checks if the user has accidently changed one cell on the spreadsheet that they shouldn't have. If the oldValue is not undefined, then 
 * it places the previous value back into the active range.
 * 
 * @param {Sheet}          sheet    : The Item Search sheet
 * @param {Event Object}     e      : The event object
 * @param {Range}          range    : The active range
 * @param {Boolean}     isSingleRow : Whether or not a single row was editted
 * @param {Spreadsheet} spreadsheet : The active spreadsheet
 * @author Jarren Ralf
 * @OnlyCurrentDoc
 */
function undoUserMistake(sheet, e, range, isSingleRow, spreadsheet)
{
  const startTime = new Date().getTime(); // Used for the function runtime

  if (isSingleRow && e.oldValue != undefined) // Single Cell is being changed
  {
    range.setValue(e.oldValue);
    spreadsheet.toast('User Change has been undone.')
    sheet.getRange(2, 5).setValue((new Date().getTime() - startTime)/1000 + " seconds");
  }
}

/**
* This function checks if the user has pressed delete on a certain cell or not, returning true if they have.
*
* @param {String or Undefined} value : An inputed string or undefined
* @return {Boolean} Returns a boolean reporting whether the event object new value is undefined or not.
* @author Jarren Ralf
* @OnlyCurrentDoc
*/
function userHasPressedDelete(value)
{
  return value === undefined;
}