/**
 * This function handles the on edit events in this spreadsheet pertaining to the Item Search sheet only (all other sheets will be protected).
 * This function is looking for the user searching for items and it is making appropriate changes to the data when a user deletes items from their order.
 * 
 * @param {Event Object} e : The event object
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

  if (sheet.getSheetName() === 'Item Search' && isSingleColumn)
    if (row == 1 && col == 1 && (rowEnd == null || rowEnd == 2 || isSingleRow))
      search(e, spreadsheet, sheet);
    else if (row == 2 && col == 6) // Submission Checkbox
      checkForOrderSubmission(range);
    else if (row > 4) // If the body of the item Search is being edited
      if (col == 8) // Items are being selected in the description column
        deleteItemsFromOrder(sheet, range, range.getValue(), row, isSingleRow, spreadsheet);
      else if (col == 1 || col == 4 || col == 7) // The SKU, UoM, or the Descriptions - Categories - Unit of Measure - SKU # column are being edited (The user is not suppose to edit these fields)
        undoUserMistake(sheet, e, range, isSingleRow, spreadsheet)
}

/**
 * This function identifies all of the cells that the user has selected and moves those items to the order portion of the Item Search sheet.
 * 
 * @author Jarren Ralf
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
    const row = (isNotBlank(sheet.getSheetValues(5, 4, 1, 1)[0][0])) ? 
      Math.max(getLastRowSpecial(sheet.getSheetValues(1, 4, sheet.getMaxRows(), 1)), // SKU column
               getLastRowSpecial(sheet.getSheetValues(1, 8, sheet.getMaxRows(), 1))) // Description column
      + 1: 5;
    sheet.getRange(row, 3, numItems, 6).setNumberFormat('@').setValues(itemValues.map(item => {
      splitDescription = item[0].split(' - ');
      sku = splitDescription.pop();
      uom = splitDescription.pop();
      splitDescription.pop();
      return ['D', sku, 0, '', uom, splitDescription]
    })).offset(0, 2, 1, 1).activate(); // Move to the quantity column
  }
  else
    SpreadsheetApp.getUi().alert('Please select an item from the list.');

  sheet.getRange(2, 7).setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * This function retrieves the items on the Recently Created and places them on the Item Search sheet.
 * 
 * @author Jarren Ralf
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
    .offset(-3, 7, 1, 1).setValue("Items displayed in order of newest to oldest.") // Tell user items are sorted from newest to oldest
    .offset(0, -1).setValue((new Date().getTime() - startTime)/1000 + " seconds"); // Function runtime
  spreadsheet.toast('PNT\'s most recently created items are being displayed.');
}

/**
 * This function checks if the customer has checked or unchecked their submission button. In both cases, this will lead to the PNT Order Processor spreadsheet to send the appropriate 
 * emails to the customer and PNT employees.
 * 
 * @param {Range} range : The range of the checkbox.
 * @author Jarren Ralf
 */
function checkForOrderSubmission(range)
{
  const startTime = new Date().getTime(); // Used for the function runtime
  range.offset(0, 1).setValue((new Date().getTime() - startTime)/1000 + " seconds").offset(2, -5).setValue(startTime); // Place the timestamp in one of the hidden cells
  SpreadsheetApp.flush(); // Force the change on the spreadsheet first
  SpreadsheetApp.getUi().alert((range.isChecked()) ? 'Your order has been submitted.\n\nThank You!' : 'You have cancelled your order.\n\nYou may make changes and re-submit by selecting the checkbox again.')
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
 */
function deleteItemsFromOrder(sheet, range, value, row, isSingleRow, spreadsheet)
{
  const startTime = new Date().getTime(); // Used for the function runtime
  spreadsheet.toast('Checking for possible lines to delete...')
  const numRows = Math.max(getLastRowSpecial(sheet.getSheetValues(1, 4, sheet.getMaxRows(), 1)), getLastRowSpecial(sheet.getSheetValues(1, 8, sheet.getMaxRows(), 1))) - row + 1;

  if (numRows > 0)
  {
    const itemsOrderedRange = sheet.getRange(row, 3, numRows, 6);
    
    if (isSingleRow)
    {
      if (!Boolean(value)) // Was a single cell editted?, is the value blank? or is the quantity zero?
      {
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
      const orderedItems = itemsOrderedRange.getValues().filter(description => isNotBlank(description[5])); // Find and remove the blank descriptions
      itemsOrderedRange.clearContent();
      
      if (orderedItems.length > 0)
        itemsOrderedRange.offset(0, 0, orderedItems.length, 6).setValues(orderedItems); // Move the items up to fill in the gaps 

      spreadsheet.toast('Deleting Complete.')
    }
    else
      spreadsheet.toast('Nothing Deleted.')
  }
  else
    spreadsheet.toast('Nothing Deleted.')

  sheet.getRange(2, 7).setValue((new Date().getTime() - startTime)/1000 + " seconds");
}

/**
 * Gets the last row number based on a selected column range values
 *
 * @param {array} range : takes a 2d array of a single column's values
 * @returns {number} : the last row number with a value. 
 */ 
function getLastRowSpecial(range)
{
  for (var row = 0, rowNum = 0, blank = false; row < range.length; row++)
  {
    if (isBlank(range[row][0]) && !blank)
    {
      rowNum = row;
      blank = true;
    }
    else if (isNotBlank(range[row][0]))
      blank = false;
  }
  return rowNum;
}

/**
 * This function checks if the given string is blank or not.
 * 
 * @param {String} str : The given string.
 * @returns {Boolean} Whether the given string is blank or not.
 * @author Jarren Ralf
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
        .offset(-3, 6, 1, 1).setValue("No results found.\nPlease try again.");
    else
      sheet.getRange('A5').activate() // Move the user to the top of the search items
        .offset(0, 0, sheet.getMaxRows() - 4).clearContent()
        .offset(0, 0, numItems).setValues(output) 
        .offset(-3, 7, 1, 1).setValue((numItems !== 1) ? numItems + " results found." : numItems + " result found.");

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
      .offset(-3, 7, 1, 1).setValue("Items displayed in order of newest to oldest.")
    spreadsheet.toast('PNT\'s most recently created items are being displayed.')
  }
  else
  {
    sheet.getRange(5, 1, sheet.getMaxRows() - 4).clearContent() // Clear content 
      .offset(-3, 7, 1, 1).setValue("Invalid search.\nPlease try again.");
    spreadsheet.toast('Invalid Search.');
  }

  sheet.getRange(2, 7).setValue((new Date().getTime() - startTime)/1000 + " seconds");
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
 */
function undoUserMistake(sheet, e, range, isSingleRow, spreadsheet)
{
  const startTime = new Date().getTime(); // Used for the function runtime

  if (isSingleRow && e.oldValue != undefined) // Single Cell is being changed
  {
    range.setValue(e.oldValue);
    spreadsheet.toast('User Change has been undone.')
    sheet.getRange(2, 7).setValue((new Date().getTime() - startTime)/1000 + " seconds");
  }
}

/**
* This function checks if the user has pressed delete on a certain cell or not, returning true if they have.
*
* @param {String or Undefined} value : An inputed string or undefined
* @return {Boolean} Returns a boolean reporting whether the event object new value is undefined or not.
* @author Jarren Ralf
*/
function userHasPressedDelete(value)
{
  return value === undefined;
}