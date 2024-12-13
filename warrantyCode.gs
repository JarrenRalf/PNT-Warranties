/**
 * This function handles the onChange trigger event. It checks when rows are added to the Status page, and deletes them all except for one.
 * 
 * @param {Event} e : The event object 
 */
function onChange(e)
{
  if (e.changeType === 'INSERT_ROW')
    manageStatusPage(e)
}

/**
 * This function handles the onEdit trigger event. It checks when the Status Page or the Repair Form are edited and saves those changes in 
 * the All_Active_Warrenties page appropriately. It also populates the repair form with current or completed orders.
 * 
 * @param {Event} e : The event object 
 */
function installedOnEdit(e)
{
  var range = e.range; 
  var row = range.rowStart;
  var col = range.columnStart;
  var spreadsheet = e.source;
  var sheet = spreadsheet.getActiveSheet();
  var sheetName = sheet.getSheetName();

  if (sheetName === 'Repair Form') // Since there are merge cells
    updateAllActiveWarranties_RepairForm(e, range, row, col, sheet, spreadsheet)
  else if (row == range.rowEnd && col == range.columnEnd) // Single cell
  {
    if (sheetName === 'Status Page')
    {
      if (row > 2)
      {
        if (col == 2) // User wants to populate the repair form
          populateRepairForm(range, spreadsheet)
        else // User is editing the status page which actively makes changes to the All_Active_Warranties
          updateAllActiveWarranties_StatusPage(e, range, row - 3, col, sheet, spreadsheet)
      }
      else if (col === 3)
        populateRepairForm(range, spreadsheet, e.value)
    }
  }
}

/**
 * This function handles the onChange trigger event. It creates some menu items as well as refreshes the data.
 * 
 * @author Jarren
 */
function onOpen()
{
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PNT Controls')
    .addItem('Create New from Current Repair Form', 'createNew_FromRepairForm')
    .addItem('Request a New Status', 'requestNewStatus')
    .addSeparator()
    .addSubMenu(ui.createMenu('Add')
      .addItem('Employee Name', 'addEmployeeName')
      .addItem('Supplier', 'addSupplier'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Remove')
      .addItem('Employee Name', 'removeEmployeeName')
      .addItem('Supplier', 'removeSupplier'))
    // .addSeparator()
    // .addItem('Watch Help Video', 'launchVideo')
    .addToUi();

  const spreadsheet = refresh(); // Update the Repair form and the Status Page from the All_Active_Warranties page
  spreadsheet.toast('Data Refresh: Complete')
}

/**
 * If a warranty's status becomes "Sent to Parksville for Repair", then that Tag# will be added to the transfer sheet.
 * 
 * @author Jarren Ralf
 */
function addWarrantyToTransferSheet(warrantyValues, fromLocation, spreadsheet)
{
  const today = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy');
  var row = 0, numRows = 1, sheet, itemValues, url, items = warrantyValues[17];

  switch (fromLocation)
  {
    case 'Richmond':
      url = 'https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit'
      sheet = SpreadsheetApp.openByUrl(url).getSheetByName('Shipped')

      for (var i = 18; i <= 22; i++)
      {
        if (!isBlank(warrantyValues[i]))
          items += '\n' + warrantyValues[i]
      }

      itemValues = [[today, 'Warranty\nLog', 1, 'EACH', items, 'ATTN: Martin, Nate, or Noah (Repair Items)\nWarranty Tag# ' + warrantyValues[0] + '\n' + warrantyValues[1]]]
      row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, numRows, 6).setNumberFormat('@').setValues(itemValues)
      applyFullRowFormatting(sheet, row, numRows, false)
      sheet.autoResizeRows(row, 1).getRange(row, 10).setDataValidation(sheet.getRange(3, 10).getDataValidation()).offset(0, 3).setDataValidation(null)
      spreadsheet.toast('Added to the Parksville transfer sheet', 'Tag# ' + tagNum, 60)

      break;

    case 'Parksville':
      url = 'https://docs.google.com/spreadsheets/d/181NdJVJueFNLjWplRNsgNl0G-sEJVW3Oy4z9vzUFrfM/edit'
      sheet = SpreadsheetApp.openByUrl(url).getSheetByName('ItemsToRichmond')

      for (var i = 18; i <= 22; i++)
      {
        if (!isBlank(warrantyValues[i]))
          items += '\n' + warrantyValues[i]
      }

      itemValues = [[today, 'Warranty\nLog', 'EACH', items, 'ATTN: Mark\nCOMPLETED Warranty Tag# ' + warrantyValues[0] + '\n' + warrantyValues[1], 1]]
      row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, numRows, 6).setNumberFormat('@').setValues(itemValues)
      applyFullRowFormatting(sheet, row, numRows, true)
      spreadsheet.toast('Added to the ItemsToRichmond page of the Parksville transfer sheet', 'Tag# ' + tagNum, 60)
      break;

    case 'Rupert':
      url = 'https://docs.google.com/spreadsheets/d/1cK1xrtJMeMbfQHrFc_TWUwCKlYzmkov0_zuBxO55iKM/edit'
      sheet = SpreadsheetApp.openByUrl(url).getSheetByName('ItemsToRichmond')
      
      for (var i = 18; i <= 22; i++)
      {
        if (!isBlank(warrantyValues[i]))
          items += '\n' + warrantyValues[i]
      }

      itemValues = [[today, 'Warranty\nLog', 'EACH', items, 'ATTN: Martin, Nate, or Noah (Repair Items)\nWarranty Tag# ' + warrantyValues[0] + '\n' + warrantyValues[1], 1]]
      row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, numRows, 6).setNumberFormat('@').setValues(itemValues)
      applyFullRowFormatting(sheet, row, numRows, true)
      spreadsheet.toast('Added to the ItemsToRichmond page of the Rupert transfer sheet', 'Tag# ' + tagNum, 60)
      break;
  }
}

/**
 * This fucntion adds one row to the status page, which represents a new warranty
 * 
 * @author Jarren Ralf
 */
function addOneRow()
{
  const statusPage = SpreadsheetApp.getActiveSheet()
  const lastRow = statusPage.getLastRow()
  statusPage.insertRowAfter(lastRow).getRange(lastRow + 1, 1).activate()
  SpreadsheetApp.getActive().toast('Please enter a Tag#', 'New Warranty', 10)
}

/**
 * This function adds either an employee or supplier name to the appropriate data validation range.
 * 
 * @param {String} action : The action ehich the user to taking to change the data validation lists, either add or remove.
 * @param {String}  type  : The type of names that the user is trying to add to the data validation, either employee or supplier.
 * @author Jarren Ralf
 */
function add_remove(action, type)
{
  const ui = SpreadsheetApp.getUi()
  const response = ui.prompt('Type the ' + type + ' name:')

  if (response.getSelectedButton() === ui.Button.OK)
  {
    const spreadsheet = SpreadsheetApp.getActive()
    const dataValidationSheet = spreadsheet.getSheetByName('Status_Supplier_Name')
    const dataValidationRange = dataValidationSheet.getRange(1, (type === 'supplier') ? 2 : 3, dataValidationSheet.getLastRow(), 1)
    const dataValidation = dataValidationRange.getValues();
    const numDataValidation = dataValidation.length;
    const response_Proper = toProper(response.getResponseText())

    switch (action)
    {
      case 'add':

        for (var i = 0; i < numDataValidation; i++)
        {
          if (isBlank(dataValidation[i][0]))
          {
            dataValidation[i][0] = response_Proper;
            break;
          }
        }

        if (i === numDataValidation)
        {
          dataValidation.push([response_Proper])
          dataValidationRange.offset(0, 0, numDataValidation + 1, 1).setValues(dataValidation.sort((a, b) => (isBlank(a[0])) ? 1 : (isBlank(b[0])) ? -1 : (a[0] < b[0]) ? -1 : 1))
        }
        else
          dataValidationRange.setValues(dataValidation.sort((a, b) => (isBlank(a[0])) ? 1 : (isBlank(b[0])) ? -1 : (a[0] < b[0]) ? -1 : 1))

        break
      case 'remove':

        const response_LowerCase = response_Proper.toLowerCase()

        for (var i = 0; i < numDataValidation; i++)
        {
          if (dataValidation[i][0].toString().toLowerCase() === response_LowerCase)
          {
            
            break;
          }
        }

        if (i < numDataValidation)
        {
          dataValidation.splice(i, 1);
          dataValidationRange.clearContent().offset(0, 0, numDataValidation - 1, 1).setValues(dataValidation)
        }

        break;
    }

    spreadsheet.toast(response_Proper + ' has been ' + ((action === 'add') ? 'added to' : 'removed from') + ' the list.', toProper(action) + ' ' + toProper(type), 20)
  }
}

/**
 * This function adds an employee name to the data validation list.
 * 
 * @author Jarren Ralf
 */
function addEmployeeName()
{
  add_remove('add', 'employee')
}

/**
 * This function adds a supplier to the data validation list.
 * 
 * @author Jarren Ralf
 */
function addSupplier()
{
  add_remove('add', 'supplier')
}

/**
 * Apply the proper formatting to the Order, Shipped, Received, ItemsToRichmond, Manual Counts, or InfoCounts page.
 *
 * @param {Sheet}   sheet  : The current sheet that needs a formatting adjustment
 * @param {Number}   row   : The row that needs formating
 * @param {Number} numRows : The number of rows that needs formatting
 * @param {Boolean} isItemsToRichmondPage : Whether items are being added to the "ToRichmond" pages on the Parksville or Rupert spreadsheets.
 * @author Jarren Ralf
 */
function applyFullRowFormatting(sheet, row, numRows, isItemsToRichmondPage)
{
  const BLUE = '#c9daf8', GREEN = '#d9ead3', YELLOW = '#fff2cc', GREEN_DATE = '#b6d7a8';

  if (isItemsToRichmondPage)
  {
    var      borderRng = sheet.getRange(row, 1, numRows, 8);
    var  shippedColRng = sheet.getRange(row, 6, numRows   );
    var thickBorderRng = sheet.getRange(row, 6, numRows, 3);
    var backgroundColours = [...Array(numRows)].map(_ => [GREEN_DATE, 'white', 'white', 'white', 'white', 'white', 'white', 'white'])
    var numberFormats = [...Array(numRows)].map(_ => ['dd MMM yyyy', '@', '@', '@', '@', '#.#', '@', '@']);
    var horizontalAlignments = [...Array(numRows)].map(_ => ['right', 'center', 'center', 'left', 'center', 'center', 'center', 'left']);
    var wrapStrategies = [...Array(numRows)].map(_ => [...new Array(2).fill(SpreadsheetApp.WrapStrategy.OVERFLOW), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP), 
        SpreadsheetApp.WrapStrategy.CLIP, SpreadsheetApp.WrapStrategy.WRAP, SpreadsheetApp.WrapStrategy.WRAP]);
  }
  else
  {
    var         borderRng = sheet.getRange(row, 1, numRows, 11);
    var     shippedColRng = sheet.getRange(row, 9, numRows    );
    var    thickBorderRng = sheet.getRange(row, 9, numRows,  2);
    var backgroundColours = [...Array(numRows)].map(_ => [GREEN_DATE, 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white'])
    var numberFormats = [...Array(numRows)].map(e => ['dd MMM yyyy', '@', '#.#', '@', '@', '@', '#.#', '0.#', '#.#', '@', 'dd MMM yyyy']);
    var horizontalAlignments = [...Array(numRows)].map(e => ['right', ...new Array(3).fill('center'), 'left', ...new Array(6).fill('center')]);
    var wrapStrategies = [...Array(numRows)].map(e => [...new Array(3).fill(SpreadsheetApp.WrapStrategy.OVERFLOW), ...new Array(3).fill(SpreadsheetApp.WrapStrategy.WRAP),
      ...new Array(3).fill   (SpreadsheetApp.WrapStrategy.CLIP), SpreadsheetApp.WrapStrategy.WRAP, SpreadsheetApp.WrapStrategy.CLIP]);
  }

  borderRng.setFontSize(10).setFontLine('none').setFontWeight('bold').setFontStyle('normal').setFontFamily('Arial').setFontColor('black')
    .setNumberFormats(numberFormats).setHorizontalAlignments(horizontalAlignments).setWrapStrategies(wrapStrategies)
    .setBorder(true, true, true, true,  null, true, 'black', SpreadsheetApp.BorderStyle.SOLID).setBackgrounds(backgroundColours);

  thickBorderRng.setBorder(null, true, null, true, false, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setBackground(GREEN);
  shippedColRng.setBackground(YELLOW);

  if (!isItemsToRichmondPage)
    sheet.getRange(row, 7, numRows, 2).setBorder(null,  true,  null,  null,  true,  null, 'black', SpreadsheetApp.BorderStyle.SOLID).setBackground(BLUE);
}

/**
 * This function takes the given string, splits it at the chosen delimiter, and capitalizes the first letter of each perceived word.
 * 
 * @param {String} str : The given string
 * @param {String} delimiter : The delimiter that determines where to split the given string
 * @return {String} The output string with proper case
 * @author Jarren Ralf
 */
function capitalizeSubstrings(str, delimiter)
{
  var numLetters;
  var words = str.toString().split(delimiter); // Split the string at the chosen location/s based on the delimiter

  for (var word = 0, string = ''; word < words.length; word++) // Loop through all of the words in order to build the new string
  {
    numLetters = words[word].length;

    if (numLetters == 0) // The "word" is a blank string (a sentence contained 2 spaces)
      continue; // Skip this iterate
    else if (numLetters == 1) // Single character word
    {
      if (words[word][0] !== words[word][0].toUpperCase()) // If the single letter is not capitalized
        words[word] = words[word][0].toUpperCase(); // Then capitalize it
    }
    else if (numLetters == 2 && words[word].toUpperCase() === 'PO') // So that PO Box is displayed correctly
      words[word] = words[word].toUpperCase();
    else
    {
      /* If the first letter is not upper case or the second letter is not lower case, then
       * capitalize the first letter and make the rest of the word lower case.
       */
      if (words[word][0] !== words[word][0].toUpperCase() || words[word][1] !== words[word][1].toLowerCase())
        words[word] = words[word][0].toUpperCase() + words[word].substring(1).toLowerCase();
    }

    string += words[word] + delimiter; // Add a blank space at the end
  }

  string = string.slice(0, -1); // Remove the last space

  return string;
}

/**
 * This function formats the given date into 'dd MMM yyyy' eg. 09 SEP 2023.
 * 
 * @param {String} dateString : Assumed to be a string of a date.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @return {String} The formatted date string
 * @author Jarren Ralf
 */
function formatDate(dateString, spreadsheet)
{
  const timeZone = spreadsheet.getSpreadsheetTimeZone()
  const dateFormat = 'dd MMM yyyy'
  const dataType = typeof dateString;
  const month = {'Jan': 0, 'Feb': 1, 'Mar': 2, 'Apr': 3, 'May': 4, 'Jun': 5, 'Jul': 6, 'Aug': 7, 'Sep': 8, 'Oct': 9, 'Nov': 10, 'Dec': 11}
  var formattedDate, splitDateString;

  if (dataType === 'string' && !isBlank(dateString))
  {
    splitDateString = dateString.split(' ')

    if (splitDateString.length === 3)
      formattedDate = Utilities.formatDate(new Date(parseInt(splitDateString[2]), month[splitDateString[1]], parseInt(splitDateString[0])), timeZone, dateFormat); //Date
    else if (splitDateString.length === 1)
    {
      splitDateString = dateString.split('/')

      if (splitDateString.length === 3)
        formattedDate = Utilities.formatDate(new Date(parseInt(splitDateString[2]), parseInt(splitDateString[0]) - 1, parseInt(splitDateString[1])), timeZone, dateFormat); //Date
      else
        Logger.log('This format: ' + dateString + ' for an DataType: string is not currently supported in the code for formatDate()')
    }
  }
  else if (dataType === 'object') // Assumed to be date type
    formattedDate = Utilities.formatDate(dateString, timeZone, dateFormat); //Date

  return formattedDate
}

/**
 * This function creates a new warranty from the existing information on the repair form currently.
 * 
 * @author Jarren Ralf
 */
function createNew_FromRepairForm()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const repairForm = spreadsheet.getActiveSheet()

  if (repairForm.getSheetName() !== 'Repair Form')
    spreadsheet.getSheetByName('Repair Form').activate()
  else
  {
    repairForm.getRange(1, 7, 1, 2).setBackgrounds([['white', 'white']]).setBorder(false, false, true, false, false, null).setValues([['', '']])
    const warrantiesSheet = spreadsheet.getSheetByName('All_Active_Warranties');
    const statusPage = spreadsheet.getSheetByName('Status Page');
    const repairFormRange = repairForm.getRange(2, 1, repairForm.getMaxRows() - 1, repairForm.getLastColumn());
    const repairFormValues = repairFormRange.getValues();

    repairFormValues[ 0][7] = ''
    repairFormValues[ 1][7] = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy')
    repairFormValues[ 6][7] = ''
    repairFormValues[ 7][7] = ''
    repairFormValues[ 8][7] = ''
    repairFormValues[ 9][7] = ''
    repairFormValues[10][7] = ''
    repairFormValues[11][7] = ''

    repairFormValues[ 5][2] = ''
    repairFormValues[ 6][1] = ''
    repairFormValues[ 7][1] = ''
    repairFormValues[ 8][1] = ''
    repairFormValues[ 9][1] = ''
    repairFormValues[10][1] = ''
    repairFormValues[11][1] = ''

    repairFormValues[12][2] = ''
    repairFormValues[13][1] = ''
    repairFormValues[14][1] = ''
    repairFormValues[15][1] = ''

    repairFormValues[19][0] = ''
    repairFormValues[19][2] = ''
    repairFormValues[19][3] = ''
    repairFormValues[19][4] = ''

    repairFormValues[20][0] = ''
    repairFormValues[20][2] = ''
    repairFormValues[20][3] = ''
    repairFormValues[20][4] = ''

    repairFormValues[21][0] = ''
    repairFormValues[21][2] = ''
    repairFormValues[21][3] = ''
    repairFormValues[21][4] = ''

    repairFormValues[22][0] = ''
    repairFormValues[22][2] = ''
    repairFormValues[22][3] = ''
    repairFormValues[22][4] = ''

    repairFormValues[23][0] = ''
    repairFormValues[23][2] = ''
    repairFormValues[23][3] = ''
    repairFormValues[23][4] = ''

    repairFormValues[24][0] = ''
    repairFormValues[24][2] = ''
    repairFormValues[24][3] = ''
    repairFormValues[24][4] = ''

    repairFormValues[25][0] = ''
    repairFormValues[25][2] = ''
    repairFormValues[25][3] = ''
    repairFormValues[25][4] = ''

    var postCode = repairFormValues[2][2].split('  ');
    var address = postCode[0].split(', ');
    var street = '', city = '', province = '', postalCode = '';

    if (!isBlank(address[0]))
    {
      if (postCode.length == 2)
        postalCode = postCode[1]

      switch (address.length)
      {
        case 2:
          street = address[0]
          city = address[1]
          break;
        case 3:
          street = address[0]
          city = address[1]
          province = address[2]
          break;
        case 4:
          street = address[0]
          city = address[1]
          province = address[2]
          postalCode = address[3]
          break;
        default:
          street = address[0];
      }
    }

    statusPage.appendRow([
      null,                   // Tag#
      null,                   // Checkbox
      repairFormValues[1][7], // Created Date
      repairFormValues[3][7], // Created By
      repairFormValues[0][2], // Customer Name
      repairFormValues[0][4], // Company Name
      null,                   // Current Status 
      street,                 // Address
      city,                   // City
      province,               // Province
      repairFormValues[3][2], // Phone
      repairFormValues[3][4], // Email
      null,                   // Items
      repairFormValues[2][7]  // Type
    ])

    warrantiesSheet.appendRow([
      null,                   // Tag#
      repairFormValues[0][2], // Customer Name
      repairFormValues[0][4], // Company Name
      street,                 // Address
      city,                   // City
      province,               // Province
      postalCode,             // Postal Code
      repairFormValues[3][2], // Phone
      repairFormValues[3][4], // Email
      repairFormValues[1][7], // Created Date
      repairFormValues[2][7], // Type
      repairFormValues[3][7], // PNT Contact
      repairFormValues[5][7], // Supplier
      ...new Array(warrantiesSheet.getLastColumn() - 13).fill('')
    ])

    repairFormRange.setValues(repairFormValues).offset(0, 7, 1, 1).activate();
    spreadsheet.toast('Please enter a Tag#', 'New Warranty', 10);
  }
}

/**
 * This function reformats a valid phone number into (###) ###-####, unless there are too many/few digits in the number, in which case the original string is returned.
 * It handles inputs that include leading ones and pluses, as well as strings that contain or don't contain parenthesis.  
 * 
 * @param {Number} num : The given phone number
 * @return Returns a reformatted phone number
 * @author Jarren Ralf
 */
function formatPhoneNumber(num)
{
  var ph = num.toString().trim().replace(/['\])}[\s{(+-]/g, ''); // Remove any brackets, braces, parenthesis, apostrophes, dashes, plus symbols, and blank spaces

  return (ph.length === 10 && ph[0] !== '1') ? '(' + ph.substring(0, 3) + ') ' + ph.substring(3, 6) + '-' + ph.substring(6) : 
         (ph.length === 11 && ph[0] === '1') ? '(' + ph.substring(1, 4) + ') ' + ph.substring(4, 7) + '-' + ph.substring(7) : num;
}

/**
 * This function reformats a valid canadian postal code into A1A 1A1, unless there are too many/few digits in the number, in which case the original string is returned.
 * 
 * @param {Number} num : The given postal code
 * @return Returns a reformatted candian postal code
 * @author Jarren Ralf
 */
function formatPostalCode(num)
{
  var postCode = num.toString().trim().toUpperCase(); 

  return (postCode.length === 6) ? postCode.substring(0, 3) + ' ' + postCode.substring(3, 6) : postCode;
}

/**
 * This function checks if the given string is blank or not.
 * 
 * @param {String} str : The given string.
 * @return {Boolean} True if the given string is blank and false otherwise.
 * @author Jarren Ralf
 */
function isBlank(str)
{
  return str === '';
}

/**
 * This function deploys a modal dialgue box with a video that explains the features of this spreadsheet and how to use it.
 * 
 * @author Jarren Ralf
 */
function launchVideo()
{

}

/**
 * This function manages the rows added to the status page. When there is too many, it deletes all expect for one.
 * 
 * @author Jarren Ralf
 */
function manageStatusPage(e)
{
  const spreadsheet = e.source
  const sheet = spreadsheet.getActiveSheet();

  if (sheet.getSheetName() === 'Status Page')
  {
    const maxRows = sheet.getMaxRows();
    const values = sheet.getSheetValues(1, 1, maxRows, 1);
    var numRows = -1;

    for (var i = maxRows - 1; i >=3; i--)
    {
      if (isBlank(values[i][0]))
        numRows++;
      else
        break;
    }

    if (numRows > 0)
      sheet.deleteRows(i + 3, numRows)

    sheet.getRange(i + 2, 1).activate()

    spreadsheet.toast('Please enter a Tag#', 'New Warranty', 10)
  }
}

/**
 * This function populates the repair form with either a current warranty or a completed one.
 * 
 * @author Jarren Ralf
 */
function populateRepairForm(range, spreadsheet, complete)
{
  const repairForm = spreadsheet.getSheetByName('Repair Form');
  const completedDateRange = repairForm.getRange(1, 7, 1, 2);

  if (arguments.length === 3)
  {
    var index = 60;
    var warrantiesSheet = spreadsheet.getSheetByName('Completed_Warranties');
    var tagNum = complete;
    range.setValue('Completed Warranties')
    completedDateRange.setBackgrounds([['#efefef', 'white']]).setBorder(true, true, true, true, false, null)
  }
  else
  {
    var index = 0;
    var warrantiesSheet = spreadsheet.getSheetByName('All_Active_Warranties');
    var tagNum = range.uncheck().offset(0, -1).getValue().toString()
    completedDateRange.setBackgrounds([['white', 'white']]).setBorder(false, false, true, false, false, null).setValues([['', '']])
  }

  const values = warrantiesSheet.getSheetValues(2, 1, warrantiesSheet.getLastRow() - 1, warrantiesSheet.getLastColumn())
    .filter(tag => tag[index] === tagNum)[0];

  if (values == null)
    SpreadsheetApp.getUi().alert('No information was found for Tag# ' + tagNum + '.')
  else
  {
    const repairForm = spreadsheet.getSheetByName('Repair Form');
    range = repairForm.getRange(2, 1, repairForm.getMaxRows() - 1, repairForm.getMaxColumns());
    const formValues = range.getValues();

    formValues[ 0][7] = values[ 0].toString(); //Tag #
    formValues[ 0][2] = values[ 1]; //Name
    formValues[ 0][4] = values[ 2]; //Company Name
    formValues[ 2][2] = values[ 3] + ', ' + values[4] + ', ' + values[5] + '  ' + values[6]; //Address, City, Province  Postal Code
    formValues[ 3][2] = values[ 7]; //Phone
    formValues[ 3][4] = values[ 8]; //Email
    formValues[ 1][7] = formatDate(values[9], spreadsheet) //Date Received From Customer
    formValues[ 2][7] = values[10]; //Type
    formValues[ 3][7] = values[11]; //PNT Contact
    formValues[ 5][7] = values[12]; //Supplier
    formValues[ 6][7] = values[13]; //Shipped Date
    formValues[ 7][7] = values[14]; //Carrier
    formValues[ 8][7] = values[15]; //Tracking#
    formValues[ 9][7] = values[16]; //Repair Cost
    formValues[10][7] = values[17]; //Ship Cost
    formValues[11][7] = formatDate(values[18], spreadsheet) //Ship Back Date
    formValues[ 5][2] = values[19]; //Item1
    formValues[ 6][1] = values[20]; //Item2
    formValues[ 7][1] = values[21]; //Item3
    formValues[ 8][1] = values[22]; //Item4
    formValues[ 9][1] = values[23]; //Item5
    formValues[10][1] = values[24]; //Item6
    formValues[11][1] = values[25]; //Item7
    formValues[12][2] = values[26]; //Notes1
    formValues[13][1] = values[27]; //Notes2
    formValues[14][1] = values[28]; //Notes3
    formValues[15][1] = values[29]; //Notes4
    formValues[19][0] = values[30]; //Contact1
    formValues[19][2] = values[31]; //Status1
    formValues[19][3] = formatDate(values[32], spreadsheet) //Date1
    formValues[19][4] = values[33]; //Comments1
    formValues[20][0] = values[34]; //Contact2
    formValues[20][2] = values[35]; //Status2
    formValues[20][3] = formatDate(values[36], spreadsheet) //Date2
    formValues[20][4] = values[37]; //Comments2
    formValues[21][0] = values[38]; //Contact3
    formValues[21][2] = values[39]; //Status3
    formValues[21][3] = formatDate(values[40], spreadsheet) //Date3
    formValues[21][4] = values[41]; //Comments3
    formValues[22][0] = values[42]; //Contact4
    formValues[22][2] = values[43]; //Status4
    formValues[22][3] = formatDate(values[44], spreadsheet) //Date4
    formValues[22][4] = values[45]; //Comments4
    formValues[23][0] = values[46]; //Contact5
    formValues[23][2] = values[47]; //Status5
    formValues[23][3] = formatDate(values[48], spreadsheet) //Date5
    formValues[23][4] = values[49]; //Comments5
    formValues[24][0] = values[50]; //Contact6
    formValues[24][2] = values[51]; //Status6
    formValues[24][3] = formatDate(values[52], spreadsheet) //Date6
    formValues[24][4] = values[53]; //Comments6
    formValues[25][0] = values[54]; //Contact7
    formValues[25][2] = values[55]; //Status7
    formValues[25][3] = formatDate(values[56], spreadsheet) //Date7
    formValues[25][4] = values[57]; //Comments7

    if (complete != null)
      completedDateRange.setValues([['Completed Date:', formatDate(values[59], spreadsheet)]])

    range.setValues(formValues).activate()
  }
}

/**
 * This function refreshes the data on the Status Page and the Repair From from the All_Active_Warranties sheet.
 * 
 * @author Jarren Ralf
 */
function refresh()
{
  const spreadsheet = SpreadsheetApp.getActive()
  const statusPage = spreadsheet.getSheetByName('Status Page')
  const repairForm = spreadsheet.getSheetByName('Repair Form')
  const allActiveWarrantiesSheet = spreadsheet.getSheetByName('All_Active_Warranties');
  const allActiveWarrantiesValues = allActiveWarrantiesSheet.getSheetValues(2, 1, allActiveWarrantiesSheet.getLastRow() - 1, allActiveWarrantiesSheet.getLastColumn());
  const numRows_StatusPage = statusPage.getLastRow() - 2;
  const numRows_RepairForm = repairForm.getMaxRows();
  const statusPageRange = statusPage.getRange(3, 1, numRows_StatusPage, statusPage.getLastColumn());
  const repairFormRange = repairForm.getRange(1, 1, numRows_RepairForm, repairForm.getLastColumn());
  const allStatusPageData = statusPageRange.getValues();
  const repairFormData =  repairFormRange.getValues();
  var items;

  for (var i = 0; i < allActiveWarrantiesValues.length; i++)
  {
    if (allActiveWarrantiesValues[i][0] === repairFormData[1][7])
    {
      repairFormData[0][0] = '=HYPERLINK("https://pacificnetandtwine.com/", IMAGE("https://cdn.shopify.com/s/files/1/0018/7079/0771/files/logoh_180x@2x.jpg?v=1617320404"))';
      repairFormData[0][2] = 'Pacific Net & Twine Ltd. Warranty and Repair Form';
      repairFormData[0][5] = '';
      repairFormData[0][6] = '';
      repairFormData[0][7] = '';

      repairFormData[1][0] = '';
      repairFormData[1][1] = 'Customer Name:';
      repairFormData[1][2] = allActiveWarrantiesValues[i][1]; // Name
      repairFormData[1][5] = '';
      repairFormData[1][6] = 'Tag #';

      repairFormData[2][6] = 'Received from Customer:';
      repairFormData[2][7] = allActiveWarrantiesValues[i][8]; // Created Date

      repairFormData[3][1] = 'Address:';
      repairFormData[3][2] = allActiveWarrantiesValues[i][2] + ', ' + allActiveWarrantiesValues[i][3] + ', ' + allActiveWarrantiesValues[i][4] + '  ' + allActiveWarrantiesValues[i][5]; //Address
      repairFormData[3][6] = 'Type:';
      repairFormData[3][7] = allActiveWarrantiesValues[i][9]; // Type

      repairFormData[4][1] = 'Phone:';
      repairFormData[4][2] = allActiveWarrantiesValues[i][6]; // Phone
      repairFormData[4][3] = 'Email:';
      repairFormData[4][4] = allActiveWarrantiesValues[i][7]; // Email
      repairFormData[4][6] = 'PNT Contact:';
      repairFormData[4][7] = allActiveWarrantiesValues[i][10]; // PNT Contact

      repairFormData[5][0] = '';

      repairFormData[6][0] = 'Item(s):';
      repairFormData[6][2] = allActiveWarrantiesValues[i][18]; // Item # 1
      repairFormData[6][5] = ''; 
      repairFormData[6][6] = 'Supplier:'; 
      repairFormData[6][7] = allActiveWarrantiesValues[i][11]; // Supplier

      repairFormData[7][0] = '';
      repairFormData[7][1] = allActiveWarrantiesValues[i][19]; // Item # 2
      repairFormData[7][6] = 'Ship Date:'; 
      repairFormData[7][7] = allActiveWarrantiesValues[i][12]; // Ship Date

      repairFormData[8][0] = '';
      repairFormData[8][1] = allActiveWarrantiesValues[i][20]; // Item # 3
      repairFormData[8][6] = 'Carrier:'; 
      repairFormData[8][7] = allActiveWarrantiesValues[i][13]; // Carrier

      repairFormData[9][0] = '';
      repairFormData[9][1] = allActiveWarrantiesValues[i][21]; // Item # 4
      repairFormData[9][6] = 'Tracking#:'; 
      repairFormData[9][7] = allActiveWarrantiesValues[i][14]; // Tracking#

      repairFormData[10][0] = '';
      repairFormData[10][1] = allActiveWarrantiesValues[i][22]; // Item # 5
      repairFormData[10][6] = 'Repair Cost:'; 
      repairFormData[10][7] = allActiveWarrantiesValues[i][15]; // Repair Cost

      repairFormData[11][0] = '';
      repairFormData[11][1] = allActiveWarrantiesValues[i][23]; // Item # 6
      repairFormData[11][6] = 'Ship Cost:'; 
      repairFormData[11][7] = allActiveWarrantiesValues[i][16]; // Ship Cost

      repairFormData[12][0] = '';
      repairFormData[12][1] = allActiveWarrantiesValues[i][24]; // Item # 7
      repairFormData[12][6] = 'Ship Back Date:'; 
      repairFormData[12][7] = allActiveWarrantiesValues[i][17]; // Ship Back Date

      repairFormData[13][0] = 'Notes / Special Instructions:';
      repairFormData[13][2] = allActiveWarrantiesValues[i][25]; // Notes1
      
      repairFormData[14][0] = '';
      repairFormData[14][1] = allActiveWarrantiesValues[i][26]; // Notes2

      repairFormData[15][1] = allActiveWarrantiesValues[i][27]; // Notes3

      repairFormData[16][1] = allActiveWarrantiesValues[i][28]; // Notes4

      repairFormData[17][0] = '';

      repairFormData[18][0] = 'Status History:';
      repairFormData[18][2] = '';

      repairFormData[19][0] = 'PNT Contact';
      repairFormData[19][2] = 'Status';
      repairFormData[19][3] = 'Date';
      repairFormData[19][4] = 'Comments';

      repairFormData[20][0] = allActiveWarrantiesValues[i][29]; //Contact1
      repairFormData[20][2] = allActiveWarrantiesValues[i][30]; //Status1
      repairFormData[20][3] = allActiveWarrantiesValues[i][31]; //Date1
      repairFormData[20][4] = allActiveWarrantiesValues[i][32]; //Comment1

      repairFormData[21][0] = allActiveWarrantiesValues[i][33]; //Contact2
      repairFormData[21][2] = allActiveWarrantiesValues[i][34]; //Status2
      repairFormData[21][3] = allActiveWarrantiesValues[i][35]; //Date2
      repairFormData[21][4] = allActiveWarrantiesValues[i][36]; //Comment2
      
      repairFormData[22][0] = allActiveWarrantiesValues[i][37]; //Contact3
      repairFormData[22][2] = allActiveWarrantiesValues[i][38]; //Status3
      repairFormData[22][3] = allActiveWarrantiesValues[i][39]; //Date3
      repairFormData[22][4] = allActiveWarrantiesValues[i][40]; //Comment3

      repairFormData[23][0] = allActiveWarrantiesValues[i][41]; //Contact4
      repairFormData[23][2] = allActiveWarrantiesValues[i][42]; //Status4
      repairFormData[23][3] = allActiveWarrantiesValues[i][43]; //Date4
      repairFormData[23][4] = allActiveWarrantiesValues[i][44]; //Comment4

      repairFormData[24][0] = allActiveWarrantiesValues[i][45]; //Contact5
      repairFormData[24][2] = allActiveWarrantiesValues[i][46]; //Status5
      repairFormData[24][3] = allActiveWarrantiesValues[i][47]; //Date5
      repairFormData[24][4] = allActiveWarrantiesValues[i][48]; //Comment5

      repairFormData[25][0] = allActiveWarrantiesValues[i][49]; //Contact6
      repairFormData[25][2] = allActiveWarrantiesValues[i][50]; //Status6
      repairFormData[25][3] = allActiveWarrantiesValues[i][51]; //Date6
      repairFormData[25][4] = allActiveWarrantiesValues[i][52]; //Comment6

      repairFormData[26][0] = allActiveWarrantiesValues[i][53]; //Contact7
      repairFormData[26][2] = allActiveWarrantiesValues[i][54]; //Status7
      repairFormData[26][3] = allActiveWarrantiesValues[i][55]; //Date7
      repairFormData[26][4] = allActiveWarrantiesValues[i][56]; //Comment7

      var numFormats = [
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', 'dd MMM yyyy'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', 'dd MMM yyyy'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', 'dd MMM yyyy'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', '@', '@', '@', '@', '@'],
        ['@', '@', '@', 'dd MMM yyyy', '@', '@', '@', '@'],
        ['@', '@', '@', 'dd MMM yyyy', '@', '@', '@', '@'],
        ['@', '@', '@', 'dd MMM yyyy', '@', '@', '@', '@'],
        ['@', '@', '@', 'dd MMM yyyy', '@', '@', '@', '@'],
        ['@', '@', '@', 'dd MMM yyyy', '@', '@', '@', '@'],
        ['@', '@', '@', 'dd MMM yyyy', '@', '@', '@', '@'],
        ['@', '@', '@', 'dd MMM yyyy', '@', '@', '@', '@'],
      ]
    } 

    for ( var j = 0; j < numRows_StatusPage; j++)
    {
      if (allActiveWarrantiesValues[i][0] === allStatusPageData[j][0])
      {
        allStatusPageData[j][ 2] = allActiveWarrantiesValues[i][ 1]; //Name
        allStatusPageData[j][ 3] = allActiveWarrantiesValues[i][allActiveWarrantiesValues[0].length - 1]; //Current Status
        allStatusPageData[j][ 4] = allActiveWarrantiesValues[i][ 3]; //City
        allStatusPageData[j][ 5] = allActiveWarrantiesValues[i][ 4]; //Province
        allStatusPageData[j][ 6] = allActiveWarrantiesValues[i][ 6]; //Phone
        allStatusPageData[j][ 7] = allActiveWarrantiesValues[i][ 7]; //Email
        allStatusPageData[j][ 9] = allActiveWarrantiesValues[i][ 9]; //Type
        allStatusPageData[j][10] = allActiveWarrantiesValues[i][ 8]; //Created Date
        allStatusPageData[j][11] = allActiveWarrantiesValues[i][10]; //Created By

        items = [
          allActiveWarrantiesValues[i][18],
          allActiveWarrantiesValues[i][19],
          allActiveWarrantiesValues[i][20],
          allActiveWarrantiesValues[i][21],
          allActiveWarrantiesValues[i][22],
          allActiveWarrantiesValues[i][23]
        ]
        
        allStatusPageData[j][ 8] = items.filter(item => !isBlank(item)).join('\n')
        break;
      }
    }
  }
  
  if (numRows_StatusPage !== 0)
    statusPageRange.setNumberFormats(new Array(numRows_StatusPage).fill(['@', '#', 'dd MMM yyyy', '@', '@', '@', '@', '@', '@', '@', '@', '@', '@', '@'])).setValues(allStatusPageData)

  if (numRows_RepairForm !== 0)
    repairFormRange.setNumberFormats(numFormats).setValues(repairFormData)

  return spreadsheet;
}

/**
 * This function removes an employee name from the data validation list.
 * 
 * @author Jarren Ralf
 */
function removeEmployeeName()
{
  add_remove('remove', 'employee')
}

/**
 * This function removes a supplier from the data validation list.
 * 
 * @author Jarren Ralf
 */
function removeSupplier()
{
  add_remove('remove', 'supplier')
}

/**
 * This function prompts the user to enter the new status that theywould like to see added to the spreadsheet and then it sends that status to Adrian via email.
 * 
 * @author Jarren Ralf
 */
function requestNewStatus()
{
  const response = SpreadsheetApp.getUi().prompt('Type your proposed new status:').getResponseText()

  if (!isBlank(response))
  {
    const template = HtmlService.createTemplateFromFile('email')
    template.status = response;
    
    MailApp.sendEmail({
      to: "lb_blitz_allstar@hotmail.com", // "adrian@pacificnetandtwine.com",
      cc: "lb_blitz_allstar@hotmail.com",
      subject: "Proposed New Status on the PNT Warranty & Service Log Spreadsheet",
      htmlBody: template.evaluate().getContent(),
    })

    SpreadsheetApp.getActive().toast('Email set to Adrian proposing new status: ' + response + '.')
  }
}

/**
 * This function takes the given string and makes sure that each word in the string has a capitalized 
 * first letter followed by lower case.
 * 
 * @param {String} str : The given string
 * @return {String} The output string with proper case
 * @author Jarren Ralf
 */
function toProper(str)
{
  return capitalizeSubstrings(capitalizeSubstrings(str, '-'), ' ');
}

/**
 * This function updates the All_Active_Warranties when the repair form is edited. 
 * 
 * @author Jarren Ralf
 */
function updateAllActiveWarranties_RepairForm(e, range, row, col, repairForm, spreadsheet)
{
  const isComplete = repairForm.getSheetValues(1, 7, 1, 1)[0][0].toString() === 'Completed Date:'

  if (isComplete)
  {
    var val = (e.oldValue === undefined) ? '' : e.oldValue;
    range.setValue(val)
    SpreadsheetApp.getUi().alert('You can\'t make changes to a completed warranty.')
  }
  else
  {
    const isTagNumberEdited = (row === 2 && col === 8);
    const tagNum = (!isTagNumberEdited) ? repairForm.getSheetValues(2, 8, 1, 1)[0][0].toString() : (e.oldValue !== undefined) ? e.oldValue.toString() : range.getValue().toString();
    const allActiveWarrantiesSheet = spreadsheet.getSheetByName('All_Active_Warranties');
    const numRows = allActiveWarrantiesSheet.getLastRow() - 1
    const tagNumbers = allActiveWarrantiesSheet.getSheetValues(2, 1, numRows, 1)
    var newValue = range.getValue();

    if (isTagNumberEdited)
    {
      if (isBlank(newValue))
      {
        range.setValue(tagNum)
        SpreadsheetApp.getUi().alert('You can\'t delete the tag number.')
        return;
      }
      else if (e.oldValue === undefined)
      {
        for (var i = 0; i < numRows; i++)
        {
          if (isBlank(tagNumbers[i][0]))
          {
            allActiveWarrantiesSheet.getRange(i + 2, 1).setNumberFormat('@').setValue(tagNum)
            break;
          }
        }

        const statusPage = spreadsheet.getSheetByName('Status Page')
        const tags = statusPage.getSheetValues(3, 1, statusPage.getLastRow() - 2, 1)

        for (i = 0; i < tags.length; i++)
        {
          if (isBlank(tags[i][0]))
          {
            statusPage.getRange(i + 3, 1).setValue(tagNum)
            break;
          }
        }
        return;
      }
    }

    for (var i = 0; i < tagNumbers.length; i++)
    {
      if (tagNumbers[i][0] === tagNum)
      {
        var rowIndex = i + 2;
        break;
      }
    }

    if (i === tagNumbers.length && !isTagNumberEdited)
      SpreadsheetApp.getUi().alert('No information was found for Tag# ' + tagNum + '.')
    else
    {
      const statusPage = spreadsheet.getSheetByName('Status Page');
      const nRows = statusPage.getLastRow() - 2;
      const statusPageRange = statusPage.getRange(3, 1, nRows, statusPage.getLastColumn());
      const allStatusPageData = statusPageRange.getValues()
      var colIndex = null;

      for (var j = 0; j < nRows; j++)
        if (allStatusPageData[j][0] === tagNum)
          break;

      switch (row)
      {
        case 1:

          switch (col)
          {
            case 1:
              range.setFormula('=HYPERLINK("https://pacificnetandtwine.com/", IMAGE("https://cdn.shopify.com/s/files/1/0018/7079/0771/files/logoh_180x@2x.jpg?v=1617320404"))');
              break;
            case 3:
              range.setValue('Pacific Net & Twine Ltd. Warranty and Repair Form')
              break;
            case 7:
              (e.oldValue === 'Completed Date:') ? range.setValue('Completed Date:') : range.setValue('')
              break;
            case 8:
              (e.oldValue !== undefined) ? range.setValue(e.oldValue) : range.setValue('')
              break;
            case 2:
            case 4:
            case 5:
            case 6:
              range.setValue('')
              break;
          }
          SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
          break;
        case 2:

          switch (col)
          {
            case 2:
              range.setValue('Customer Name:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 3:
              allStatusPageData[j][4] = toProper(newValue);
              range.setValue(allStatusPageData[j][4])
              colIndex = 2;
              break;
            case 4:
              range.setValue('Company Name:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 5:
              allStatusPageData[j][5] = toProper(newValue);
              range.setValue(allStatusPageData[j][5])
              colIndex = 3;
              break;
            case 7:
              range.setValue('Tag #')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 8:

              if (e.oldValue !== undefined)
              {
                const completedWarrantiesSheet = spreadsheet.getSheetByName('Completed_Warranties')
                var allTagNumbers = [tagNumbers, statusPage.getSheetValues(3, 1, statusPage.getLastRow() - 2, 1), 
                                     completedWarrantiesSheet.getSheetValues(2, 1, completedWarrantiesSheet.getLastRow() - 1, 1)].flat(2)

                if (allTagNumbers.includes(newValue))
                {
                  range.setValue(e.oldValue)
                  SpreadsheetApp.getUi().alert('Tag# ' + newValue + ' already exists. Please choose a unique number.')

                  return;
                }
                else
                {
                  allStatusPageData[j][0] = newValue.toString().toUpperCase();
                  range.setValue(allStatusPageData[j][0])
                  colIndex = 1;
                }
              }
              else
              {
                const completedWarrantiesSheet = spreadsheet.getSheetByName('Completed_Warranties')
                var allTagNumbers = [tagNumbers, statusPage.getSheetValues(3, 1, statusPage.getLastRow() - 2, 1), 
                                     completedWarrantiesSheet.getSheetValues(2, 1, completedWarrantiesSheet.getLastRow() - 1, 1)].flat(2)

                if (allTagNumbers.includes(newValue))
                {
                  range.setValue('')
                  newValue = '';
                  SpreadsheetApp.getUi().alert('Tag# ' + newValue + ' already exists. Please choose a unique number.')

                  return;
                }
                else
                {
                  allStatusPageData[j][0] = newValue.toString().toUpperCase();
                  range.setValue(allStatusPageData[j][0])
                  colIndex = 1;
                }
              }

              break;
            case 1:
            case 6:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 3:

          switch (col)
          {
            case 7:
              range.setValue('Received from Customer:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 8:
              allStatusPageData[j][2] = newValue; // Created by date
              colIndex = 10;
              break;
            case 1:
            case 2:
            case 3:
            case 4:
            case 5:
            case 6:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 4:

          switch (col)
          {
            case 2:
              range.setValue('Address:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 3:

              const addressRange = allActiveWarrantiesSheet.getRange(rowIndex, 4, 1, 4)
              var postCode = newValue.split('  ');
              var address = postCode[0].split(', ');

              if (!isBlank(address[0]))
              {
                const addressValues = [['', '', '', '']]

                if (postCode.length == 2)
                {
                  addressValues[0][3] = formatPostalCode(postCode[1])
                  address.push(addressValues[0][3])
                }

                switch (address.length)
                {
                  case 1:
                    addressValues[0][0] = toProper(address[0]);
                    allStatusPageData[j][7] = toProper(address[0]);
                    allStatusPageData[j][8] = '';
                    allStatusPageData[j][9] = '';
                    range.setValue(addressValues[0][0] + ', ,  ')
                    break;
                  case 2:
                    addressValues[0][0] = toProper(address[0]);
                    addressValues[0][1] = toProper(address[1]);
                    allStatusPageData[j][7] = toProper(address[0]);
                    allStatusPageData[j][8] = toProper(address[1]);
                    allStatusPageData[j][9] = '';
                    range.setValue(addressValues[0][0] + ', ' + addressValues[0][1] + ',  ')
                    break;
                  case 3:
                    addressValues[0][0] = toProper(address[0]);
                    addressValues[0][1] = toProper(address[1]);
                    addressValues[0][2] = address[2].toUpperCase();
                    allStatusPageData[j][7] = toProper(address[0]);
                    allStatusPageData[j][8] = toProper(address[1]);
                    allStatusPageData[j][9] = toProper(address[2]);
                    range.setValue(addressValues[0][0] + ', ' + addressValues[0][1] + ', ' + addressValues[0][2] + '  ')
                    break;
                  case 4:
                    addressValues[0][0] = toProper(address[0]);
                    addressValues[0][1] = toProper(address[1]);
                    addressValues[0][2] = address[2].toUpperCase();
                    allStatusPageData[j][ 7] = toProper(address[0]);
                    allStatusPageData[j][ 8] = toProper(address[1]);
                    allStatusPageData[j][ 9] = toProper(address[2]);
                    allStatusPageData[j][10] = addressValues[0][3];
                    range.setValue(addressValues[0][0] + ', ' + addressValues[0][1] + ', ' + addressValues[0][2] + '  ' + addressValues[0][3])
                    break;
                  default:
                    addressValues[0][0] = toProper(address[0]);
                    allStatusPageData[j][7] = toProper(address[0]);
                    allStatusPageData[j][8] = '';
                    allStatusPageData[j][9] = '';
                    range.setValue(addressValues[0][0] + ', ,  ')
                }

                addressRange.setNumberFormat('@').setValues(addressValues)
              }
              else
                addressRange.setValue('')

              if (j !== nRows)
                statusPageRange.setNumberFormats(new Array(nRows).fill(['@', '#', 'dd MMM yyyy', '@', '@', '@', '@', '@', '@', '@', '@', '@', '@', '@'])).setValues(allStatusPageData)

              return;
            case 7:
              range.setValue('Type:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 8:
              allStatusPageData[j][13] = toProper(newValue);
              range.setValue(allStatusPageData[j][13])
              colIndex = 11;
              break;
            case 1:
            case 4:
            case 5:
            case 6:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 5:

          switch (col)
          {
            case 2:
              range.setValue('Phone:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 3:
              allStatusPageData[j][10] = formatPhoneNumber(newValue);
              range.setValue(allStatusPageData[j][10])
              colIndex = 8;
              break;
            case 4:
              range.setValue('Email:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 5:
              allStatusPageData[j][11] = newValue.toString().toLowerCase();
              range.setValue(allStatusPageData[j][11])
              colIndex = 9;
              break;
            case 7:
              range.setValue('PNT Contact:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 8:
              allStatusPageData[j][3] = toProper(newValue);
              range.setValue(allStatusPageData[j][3])
              colIndex = 12;
              manageDataValidation('Name', newValue, spreadsheet);
              break;
            case 1:
            case 6:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 7:

          switch (col)
          {
            case 1:
              range.setValue('Item(s):')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 3:
              colIndex = 20;
              const items = repairForm.getSheetValues(8, 2, 5, 1).flat().filter(u => !isBlank(u))

              if (!isBlank(newValue))
                items.unshift(newValue)

              allStatusPageData[j][12] = items.join('\n');
              break;
            case 7:
              range.setValue('Supplier:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 8:
              newValue = toProper(newValue)
              range.setValue(newValue)
              colIndex = 13;
              manageDataValidation('Supplier', newValue, spreadsheet);
              break;
            case 2:
            case 4:
            case 5:
            case 6:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 8:

          switch (col)
          {
            case 2:
              colIndex = 21;
              const items = repairForm.getSheetValues(7, 2, 7, 2).flat().filter(u => !isBlank(u) && u !== 'Item(s):')
              allStatusPageData[j][12] = items.join('\n');
              break;
            case 7:
              range.setValue('Ship Date:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 8:
              colIndex = 14;
              break;
            case 1:
            case 3:
            case 4:
            case 5:
            case 6:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 9:

          switch (col)
          {
            case 2:
              colIndex = 22;
              const items = repairForm.getSheetValues(7, 2, 7, 2).flat().filter(u => !isBlank(u) && u !== 'Item(s):')
              allStatusPageData[j][12] = items.join('\n');
              break;
            case 7:
              range.setValue('Carrier:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 8:
              newValue = toProper(newValue)
              range.setValue(newValue)
              colIndex = 15;
              break;
            case 1:
            case 3:
            case 4:
            case 5:
            case 6:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 10:

          switch (col)
          {
            case 2:
              colIndex = 23;
              const items = repairForm.getSheetValues(7, 2, 7, 2).flat().filter(u => !isBlank(u) && u !== 'Item(s):')
              allStatusPageData[j][12] = items.join('\n');
              break;
            case 7:
              range.setValue('Tracking#:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 8:
              newValue = newValue.toString().toUpperCase()
              range.setValue(newValue)
              colIndex = 16;
              break;
            case 1:
            case 3:
            case 4:
            case 5:
            case 6:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 11:

          switch (col)
          {
            case 2:
              colIndex = 24;
              const items = repairForm.getSheetValues(7, 2, 7, 2).flat().filter(u => !isBlank(u) && u !== 'Item(s):')
              allStatusPageData[j][12] = items.join('\n');
              break;
            case 7:
              range.setValue('Repair Cost:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 8:
              colIndex = 17;
              break;
            case 1:
            case 3:
            case 4:
            case 5:
            case 6:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 12:

          switch (col)
          {
            case 2:
              colIndex = 25;
              const items = repairForm.getSheetValues(7, 2, 7, 2).flat().filter(u => !isBlank(u) && u !== 'Item(s):')
              allStatusPageData[j][12] = items.join('\n');
              break;
            case 7:
              range.setValue('Ship Cost:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 8:
              colIndex = 18;
              break;
            case 1:
            case 3:
            case 4:
            case 5:
            case 6:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 13:
          
          switch (col)
          {
            case 2:
              colIndex = 26;
              const items = repairForm.getSheetValues(7, 2, 7, 2).flat().filter(u => !isBlank(u) && u !== 'Item(s):')
              allStatusPageData[j][12] = items.join('\n');
              break;
            case 7:
              range.setValue('Ship Back Date:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 8:
              colIndex = 19;
              break;
            case 1:
            case 3:
            case 4:
            case 5:
            case 6:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 14:

          switch (col)
          {
            case 1:
              range.setValue('Notes / Special Instructions:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 3:
              colIndex = 27;
              break;
            case 2:
            case 4:
            case 5:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 15:

          switch (col)
          {          
            case 2:
              colIndex = 28;
              break;
            case 1:
            case 3:
            case 4:
            case 5:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 16:

          switch (col)
          {          
            case 2:
              colIndex = 29;
              break;
            case 1:
            case 3:
            case 4:
            case 5:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 17:

          switch (col)
          {          
            case 2:
              colIndex = 30;
              break;
            case 1:
            case 3:
            case 4:
            case 5:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 19:

          switch (col)
          {
            case 1:
              range.setValue('Status History:')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 2:
            case 3:
            case 4:
            case 5:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 20:

          switch (col)
          {
            case 1:
              range.setValue('PNT Contact')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 3:
              range.setValue('Status')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 4:
              range.setValue('Date')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 5:
              range.setValue('Comments')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
            case 2:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 21:

          switch (col)
          {
            case 1:
              colIndex = 31
              newValue = toProper(newValue)
              range.setValue(newValue)
              manageDataValidation('Name', newValue, spreadsheet);
              break;
            case 3:
              colIndex = 32
              var formattedDate = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy')
              const statusHistory = repairForm.getSheetValues(21, 3, 7, 1).flat().filter(s => !isBlank(s))
              repairForm.getRange(row, col + 1).setValue(formattedDate)
              allStatusPageData[j][6] = statusHistory.pop(); // Current Status

              manageDataValidation('Status', newValue, spreadsheet);
              manageStatusChange_repairForm(spreadsheet, repairForm, allActiveWarrantiesSheet, statusPage, allStatusPageData, j,
                                            newValue, statusHistory, tagNum, formattedDate, range, rowIndex, numRows, colIndex)
              break;
            case 4:
              colIndex = 33
              break;
            case 5:
              colIndex = 34
              break;
            case 2:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 22:

          switch (col)
          {
            case 1:
              colIndex = 35
              newValue = toProper(newValue)
              range.setValue(newValue)
              manageDataValidation('Name', newValue, spreadsheet);
              break;
            case 3:
              colIndex = 36
              var formattedDate = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy')
              const statusHistory = repairForm.getSheetValues(21, 3, 7, 1).flat().filter(s => !isBlank(s))
              repairForm.getRange(row, col + 1).setValue(formattedDate)
              allStatusPageData[j][6] = statusHistory.pop(); // Current Status

              manageDataValidation('Status', newValue, spreadsheet);
              manageStatusChange_repairForm(spreadsheet, repairForm, allActiveWarrantiesSheet, statusPage, allStatusPageData, j,
                                            newValue, statusHistory, tagNum, formattedDate, range, rowIndex, numRows, colIndex)
              break;
            case 4:
              colIndex = 37
              break;
            case 5:
              colIndex = 38
              break;
            case 2:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 23:

          switch (col)
          {
            case 1:
              colIndex = 39
              newValue = toProper(newValue)
              range.setValue(newValue)
              manageDataValidation('Name', newValue, spreadsheet);
              break;
            case 3:
              colIndex = 40
              var formattedDate = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy')
              const statusHistory = repairForm.getSheetValues(21, 3, 7, 1).flat().filter(s => !isBlank(s))
              repairForm.getRange(row, col + 1).setValue(formattedDate)
              allStatusPageData[j][6] = statusHistory.pop(); // Current Status

              manageDataValidation('Status', newValue, spreadsheet);
              manageStatusChange_repairForm(spreadsheet, repairForm, allActiveWarrantiesSheet, statusPage, allStatusPageData, j,
                                            newValue, statusHistory, tagNum, formattedDate, range, rowIndex, numRows, colIndex)
              break;
            case 4:
              colIndex = 41
              break;
            case 5:
              colIndex = 42
              break;
            case 2:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 24:

          switch (col)
          {
            case 1:
              colIndex = 43
              newValue = toProper(newValue)
              range.setValue(newValue)
              manageDataValidation('Name', newValue, spreadsheet);
              break;
            case 3:
              colIndex = 44
              var formattedDate = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy')
              const statusHistory = repairForm.getSheetValues(21, 3, 7, 1).flat().filter(s => !isBlank(s))
              repairForm.getRange(row, col + 1).setValue(formattedDate)
              allStatusPageData[j][6] = statusHistory.pop(); // Current Status

              manageDataValidation('Status', newValue, spreadsheet);
              manageStatusChange_repairForm(spreadsheet, repairForm, allActiveWarrantiesSheet, statusPage, allStatusPageData, j,
                                            newValue, statusHistory, tagNum, formattedDate, range, rowIndex, numRows, colIndex)
              break;
            case 4:
              colIndex = 45
              break;
            case 5:
              colIndex = 46
              break;
            case 2:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 25:

          switch (col)
          {
            case 1:
              colIndex = 47
              newValue = toProper(newValue)
              range.setValue(newValue)
              manageDataValidation('Name', newValue, spreadsheet);
              break;
            case 3:
              colIndex = 48
              var formattedDate = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy')
              const statusHistory = repairForm.getSheetValues(21, 3, 7, 1).flat().filter(s => !isBlank(s))
              repairForm.getRange(row, col + 1).setValue(formattedDate)
              allStatusPageData[j][6] = statusHistory.pop(); // Current Status

              manageDataValidation('Status', newValue, spreadsheet);
              manageStatusChange_repairForm(spreadsheet, repairForm, allActiveWarrantiesSheet, statusPage, allStatusPageData, j,
                                            newValue, statusHistory, tagNum, formattedDate, range, rowIndex, numRows, colIndex)
              break;
            case 4:
              colIndex = 49
              break;
            case 5:
              colIndex = 50
              break;
            case 2:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 26:

          switch (col)
          {
            case 1:
              colIndex = 51
              newValue = toProper(newValue)
              range.setValue(newValue)
              manageDataValidation('Name', newValue, spreadsheet);
              break;
            case 3:
              colIndex = 52
              var formattedDate = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy')
              const statusHistory = repairForm.getSheetValues(21, 3, 7, 1).flat().filter(s => !isBlank(s))
              repairForm.getRange(row, col + 1).setValue(formattedDate)
              allStatusPageData[j][6] = statusHistory.pop(); // Current Status

              manageDataValidation('Status', newValue, spreadsheet);
              manageStatusChange_repairForm(spreadsheet, repairForm, allActiveWarrantiesSheet, statusPage, allStatusPageData, j,
                                            newValue, statusHistory, tagNum, formattedDate, range, rowIndex, numRows, colIndex)
              break;
            case 4:
              colIndex = 53
              break;
            case 5:
              colIndex = 54
              break;
            case 2:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }

          break;
        case 27:

          switch (col)
          {
            case 1:
              colIndex = 55
              newValue = toProper(newValue)
              range.setValue(newValue)
              manageDataValidation('Name', newValue, spreadsheet);
              break;
            case 3:
              colIndex = 56
              var formattedDate = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy')
              const statusHistory = repairForm.getSheetValues(21, 3, 7, 1).flat().filter(s => !isBlank(s))
              repairForm.getRange(row, col + 1).setValue(formattedDate)
              allStatusPageData[j][6] = statusHistory.pop(); // Current Status

              manageDataValidation('Status', newValue, spreadsheet);
              manageStatusChange_repairForm(spreadsheet, repairForm, allActiveWarrantiesSheet, statusPage, allStatusPageData, j,
                                            newValue, statusHistory, tagNum, formattedDate, range, rowIndex, numRows, colIndex)
              break;
            case 4:
              colIndex = 57
              break;
            case 5:
              colIndex = 58
              break;
            case 2:
            case 6:
            case 7:
            case 8:
              range.setValue('')
              SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
              break;
          }
          
          break;
        case 6:
        case 18:
          range.setValue('')
          SpreadsheetApp.getUi().alert('You can\'t edit a protected range.')
          break;
      }

      if (j !== nRows)
        statusPageRange.setNumberFormats(new Array(nRows).fill(['@', '#', 'dd MMM yyyy', '@', '@', '@', '@', '@', '@', '@', '@', '@', '@', '@'])).setValues(allStatusPageData)

      if (colIndex != null)
        allActiveWarrantiesSheet.getRange(rowIndex, colIndex).setValue(newValue)
    }
  }
}

/**
 * This function manages the change of 3 different types of data validation on the Status and Repair page, namely Status, Supplier, and Employee Names.
 * 
 * @param {String}       category   : The category of the data validation that id being edited.
 * @param {String}       element    : The element that may be a new addition to the current data validation.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function manageDataValidation(category, element, spreadsheet)
{
  const sheet = spreadsheet.getSheetByName('Status_Supplier_Name');
  const lastRow = sheet.getLastRow()
  const values = sheet.getSheetValues(1, 1, lastRow, 3);
  var col, element_Not_Found = true;

  switch (category)
  {
    case 'Status':
      col = 0;
      break;
    case 'Supplier':
      col = 1;
      element = toProper(element);
      break;
    case 'Name':
      col = 2;
      element = toProper(element);
      break;
  }

  for (var row = 0; row < lastRow; row++)
  {

    if (isBlank(values[row][col]))
    {
      values[row][col] = element
      sheet.getRange(1, 1, lastRow, 3).setValues(values)
      break;
    }
    else if (element === values[row][col])
    {
      element_Not_Found = false;
      break;
    }
  }

  if (element_Not_Found && row === lastRow) // The element was not found after looking through the whole list
  {
    values.push(['', '', ''])
    values[row][col] = element
    sheet.getRange(1, 1, lastRow + 1, 3).setValues(values)
  }
}

/**
 * This function manages the change of status on the repair form.
 * 
 * @param {Spreadsheet}      ss            : The active spreadsheet
 * @param {Sheet}        repairForm        : The Repair Form sheet.
 * @param {Sheet} allActiveWarrantiesSheet : The All_Active_Warranties sheet.
 * @param {Sheet}        statusPage        : The Status Page sheet.
 * @param {Object[][]} allStatusPageData   : All of the values displayed on the status page.
 * @param {Number}           j             : The index of the allStatusPageData for the row pertaining to the current order.
 * @param {String}         status          : The status that has just been changed.
 * @param {String[]}    statusHistory      : The list of statuses for the current warranty/repair order.
 * @param {String}         tagNum          : The tag number unique to the current warranty/repair order. 
 * @param {Date}        completedDate      : The formatted date for today. 
 * @param {Range}         editRange        : The range from the onEdit event object
 * @param {Number}          row            : The row number in the All_Active_Warranties sheet that the current warranty/repair order belongs to.
 * @param {Number}        numRows          : The current number of orders.
 * @param {Number}        colIndex         : The column number on the All_Active_Warranties sheet that corresponds to the current change on the repair form.
 * @author Jarren Ralf
 */
function manageStatusChange_repairForm(ss, repairForm, allActiveWarrantiesSheet, statusPage, allStatusPageData, j, status, statusHistory, tagNum, completedDate, editRange, row, numRows, colIndex)
{
  const range = allActiveWarrantiesSheet.getRange(row, 1, 1, allActiveWarrantiesSheet.getLastColumn());
  const values = range.getValues()[0]
  const nameRange = editRange.offset(0, -2);
  var name = nameRange.getValue(), comments = '';

  values[58] = allStatusPageData[j][3]; // Current Status

  if (status.split(' / ', 1)[0] === 'Complete') // Current Status is Complete
  {
    name = 'Someone';

    values.push(completedDate, 'Tag# ' + values[0] + ' - ' + values[1] + ' - ' + 
      formatDate(values[8], ss) + ' - ' + completedDate)

    const completedWarrantiesSheet = ss.getSheetByName('Completed_Warranties')
    const lastRow = completedWarrantiesSheet.getLastRow();
    const lastCol = completedWarrantiesSheet.getLastColumn();
    const completedOrders = completedWarrantiesSheet.getSheetValues(1, 1, lastRow, lastCol)
    const header = completedOrders.shift()
    const completedOrders_Sorted = Array.from(new Set(completedOrders.concat([values])
      .sort((a, b) => (a[lastCol - 1] < b[lastCol - 1]) ? 1 : -1).map(JSON.stringify)), JSON.parse)
    completedOrders_Sorted.unshift(header)              
    completedWarrantiesSheet.getRange(1, 1, lastRow + 1, lastCol).setNumberFormat('@').setValues(completedOrders_Sorted)

    allActiveWarrantiesSheet.deleteRow(row);

    const remainingWarranties_StatusPage = allStatusPageData.filter(tag => tag[0] !== tagNum);
    statusPage.deleteRow(numRows + 2).getRange(3, 1, numRows - 1, statusPage.getLastColumn()).setValues(remainingWarranties_StatusPage)

    repairForm.getRange(1, 7, 1, 2).setBackgrounds([['#efefef', 'white']]).setBorder(true, true, true, true, false, null)
      .setValues([['Completed Date:', completedDate]]) // Put the complete banner on the repair form

    return;
  }
  // else if (status === 'Sent to Parksville for Repair')
  // {
  //   name = 'Someone';
    
  //   if (statusHistory.length >= 1)
  //     var previousStatus = statusHistory.pop()

  //   if (previousStatus === 'Accepted in Richmond')
  //     addWarrantyToTransferSheet(values, 'Richmond', ss)
  //   else if (previousStatus === 'Accepted in Rupert')
  //     addWarrantyToTransferSheet(values, 'Rupert', ss)
  //   else
  //   {
  //     const ui = SpreadsheetApp.getUi()
  //     var response = ui.prompt('Adding Items to the Transfer Sheet', 
  //       'Please type the location name that the items are shipping from initially (Ignore upper or lowercase): \"rich" or \"pr\".', ui.ButtonSet.OK_CANCEL);

  //     var textResponse = response.getResponseText().toUpperCase();

  //     if (textResponse == 'RICH')
  //       addWarrantyToTransferSheet(values, 'Richmond', ss)
  //     else if (textResponse == 'PR')
  //       addWarrantyToTransferSheet(values, 'Rupert', ss)
  //     else
  //       ui.alert('Your typed response did not exactly match any of the location choices. Please Try again.')
  //   }
  // }
  // else if (status === 'Sent Back to Richmond for Distribution to Customer')
  // {
  //   name = 'Someone';
  //   addWarrantyToTransferSheet(values, 'Parksville', ss)
  // }
  else if (status === 'Accepted in Richmond' || status === 'Accepted in Parksville' || status === 'Accepted in Rupert')
  {
    var employeeName = repairForm.getSheetValues(5, 8, 1, 1)[0][0]
    name = isBlank(employeeName) ? 'Someone' : employeeName;
  }
  else if (isBlank(status))
  {
    name = '';
    completedDate = '';
    comments = ''; 
    editRange.offset(0, 1, 1, 2).clearContent();
  }
  else if (isBlank(name))
    name = 'Someone';

  nameRange.setValue(name) // Make the PNT Contact 'Someone' unless the status contains 'Accepted' in which case use the Employee name that created the order
  values[colIndex - 2] = name;
  values[colIndex    ] = completedDate
  values[colIndex + 1] = comments
  range.setValues([values])
}

/**
 * This function updates the All_Active_Warranties when the status page is edited.
 * 
 * @author Jarren Ralf
 */
function updateAllActiveWarranties_StatusPage(e, range, idx_TagNum, whichFieldToEdit, statusPage, spreadsheet)
{
  var row, isTag;
  const tagNum = (whichFieldToEdit !== 1) ? range.offset(0, 1 - whichFieldToEdit).getValue().toString() : (e.oldValue !== undefined) ? e.oldValue.toString() : range.getValue().toString();
  const allActiveWarrantiesSheet = spreadsheet.getSheetByName('All_Active_Warranties');
  const numRows = allActiveWarrantiesSheet.getLastRow() - 1;

  if (numRows > 0)
  {
    var allActiveWarranties = allActiveWarrantiesSheet.getSheetValues(2, 1, numRows, allActiveWarrantiesSheet.getLastColumn())
    var values = allActiveWarranties.filter((tag, r) => {
      isTag = tag[0] === tagNum; 

      if (isTag)
        row = r + 2; 
        
      return isTag;
    })[0];
  }
  else
    var values = null;

  if (values == null && whichFieldToEdit !== 1)
  {
    const ui = SpreadsheetApp.getUi();

    if (isBlank(tagNum))
    {
      const response = ui.prompt('Please enter a Tag#:')

      if (response.getSelectedButton() == ui.Button.OK)
      {
        const completedWarrantiesSheet = spreadsheet.getSheetByName('Completed_Warranties')
        const newTagNumber = response.getResponseText();
        const numRows_StatusPage = statusPage.getLastRow() - 2;
        const numCompletedWarranties = completedWarrantiesSheet.getLastRow() - 1;

        var tagNumbers = 
          (numRows > 0) ? 
            (numRows_StatusPage > 0) ? 
              (numCompletedWarranties > 0) ? 
                [statusPage.getSheetValues(3, 1, numRows_StatusPage, 1), completedWarrantiesSheet.getSheetValues(2, 1, numCompletedWarranties, 1)].flat(2) :
              [statusPage.getSheetValues(3, 1, numRows_StatusPage, 1)].flat(2) : 
            (numCompletedWarranties > 0) ? 
              [completedWarrantiesSheet.getSheetValues(2, 1, numCompletedWarranties, 1)].flat(2) : 
            [] :
          (numRows_StatusPage > 0) ? 
            (numCompletedWarranties > 0) ? 
              [allActiveWarrantiesSheet.getSheetValues(2, 1, numRows, 1), statusPage.getSheetValues(3, 1, numRows_StatusPage, 1), completedWarrantiesSheet.getSheetValues(2, 1, numCompletedWarranties, 1)].flat(2) :
            [allActiveWarrantiesSheet.getSheetValues(2, 1, numRows, 1), statusPage.getSheetValues(3, 1, numRows_StatusPage, 1)].flat(2) : 
          (numCompletedWarranties > 0) ? 
            [allActiveWarrantiesSheet.getSheetValues(2, 1, numRows, 1), completedWarrantiesSheet.getSheetValues(2, 1, numCompletedWarranties, 1)].flat(2) : 
          [allActiveWarrantiesSheet.getSheetValues(2, 1, numRows, 1)].flat(2);

        if (isBlank(newTagNumber))
        {
          ui.alert('Tag Number can\'t be blank.')
          range.setValue(e.oldValue)
          return;
        }
        else if (tagNumbers.includes(newTagNumber))
        {
          range.setValue(e.oldValue)
          SpreadsheetApp.getUi().alert('Tag# ' + newTagNumber + ' already exists. Please choose a unique number.')
          return;
        }
        else
        {
          row = allActiveWarrantiesSheet.getLastRow() + 1;
          allActiveWarrantiesSheet.getRange(row, 1).setNumberFormat('@').setValue(newTagNumber)
          range.offset(0, 1 - whichFieldToEdit).setNumberFormat('@').setValue(newTagNumber)
        }
      }
    }
    else
    {
      ui.alert('Warranty Not Found', 'No information was found for Tag# ' + tagNum + '.', ui.ButtonSet.OK)
      return;
    }
  }

  var newValue = [range.getValue()], col;

  switch (whichFieldToEdit)
  {
    case 1: // W/R Tag#

      if (e.oldValue !== undefined)
      {
        if (isBlank(newValue[0]))
        {
          range.setValue(e.oldValue.toString())
          SpreadsheetApp.getUi().alert('You can\'t delete the tag number.')
          return;
        }
        const completedWarrantiesSheet = spreadsheet.getSheetByName('Completed_Warranties')
        const tagNums = statusPage.getSheetValues(3, 1, statusPage.getLastRow() - 2, 1);
        const numCompletedWarranties = completedWarrantiesSheet.getLastRow() - 1;
        tagNums[idx_TagNum][0] = e.oldValue.toString()

        var tagNumbers = 
          (numRows > 0) ? 
            (numCompletedWarranties > 0) ? 
              [allActiveWarrantiesSheet.getSheetValues(2, 1, numRows, 1), tagNums, completedWarrantiesSheet.getSheetValues(2, 1, numCompletedWarranties, 1)].flat(2) :
            [allActiveWarrantiesSheet.getSheetValues(2, 1, numRows, 1), tagNums].flat(2) : 
          (numCompletedWarranties > 0) ? 
            [tagNums, completedWarrantiesSheet.getSheetValues(2, 1, numCompletedWarranties, 1)].flat(2) : 
          [tagNums].flat(2);
        
        if (tagNumbers.includes(newValue[0]))
        {
          range.setValue(tagNums[idx_TagNum][0])
          SpreadsheetApp.getUi().alert('Tag# ' + newValue[0] + ' already exists. Please choose a unique number.')

          return;
        }
        else
        {
          col = [1];
          newValue[0] = newValue[0].toString().toUpperCase()
        }
      }
      else
      {
        const completedWarrantiesSheet = spreadsheet.getSheetByName('Completed_Warranties')
        const tagNums = statusPage.getSheetValues(3, 1, statusPage.getLastRow() - 2, 1);
        const numCompletedWarranties = completedWarrantiesSheet.getLastRow() - 1;
        tagNums.pop()

        var tagNumbers = 
          (numRows > 0) ? 
            (numCompletedWarranties > 0) ? 
              [allActiveWarrantiesSheet.getSheetValues(2, 1, numRows, 1), tagNums, completedWarrantiesSheet.getSheetValues(2, 1, numCompletedWarranties, 1)].flat(2) :
            [allActiveWarrantiesSheet.getSheetValues(2, 1, numRows, 1), tagNums].flat(2) : 
          (numCompletedWarranties > 0) ? 
            [tagNums, completedWarrantiesSheet.getSheetValues(2, 1, numCompletedWarranties, 1)].flat(2) : 
          [tagNums].flat(2) ;

        if (tagNumbers.includes(newValue[0]))
        {
          range.setValue('')
          SpreadsheetApp.getUi().alert('Tag# ' + newValue[0] + ' already exists. Please choose a unique number.')

          return;
        }
        else
        {
          col = [1];
          row = allActiveWarrantiesSheet.getLastRow() + 1;
          newValue[0] = newValue[0].toString().toUpperCase()
        }
      }

      break;
    case 3: // Created Date
      col = [10];
      newValue[0] = formatDate(range.getDisplayValue(), spreadsheet); //Date
      break;
    case 4: // Created By
      col = [12];
      newValue[0] = toProper(newValue[0]);
      range.setValue(newValue[0])
      manageDataValidation('Name', newValue[0], spreadsheet);
      break;
    case 5: // Customer Name
      col = [2];
      newValue[0] = toProper(newValue[0])
      range.setValue(newValue[0])
      break;
    case 6:
      col = [3];
      newValue[0] = toProper(newValue[0])
      range.setValue(newValue[0])
      break;
    case 7: // Current Status

      const firstColumnForStatusData = 31;
      const range_statusHistory = allActiveWarrantiesSheet.getRange(row, firstColumnForStatusData, 1, 29);
      const statusHistory = range_statusHistory.getValues()[0];
      const timeZone = spreadsheet.getSpreadsheetTimeZone();
      const dateFormat = 'dd MMM yyyy'
      const formattedDate = Utilities.formatDate(new Date(), timeZone, dateFormat)
      var ii = 0;

      statusHistory[28] = newValue[0]; //Current Status

      for (var i = 0; i < 7; i++)
      {
        ii = i*4;

        if (isBlank(statusHistory[1 + ii]))
        {
                 values[ii + 30] = 'Someone';     // Contact
                 values[ii + 31] = newValue[0];   // Status
                 values[ii + 32] = formattedDate; // Date
          statusHistory[ii     ] = 'Someone';     // Contact
          statusHistory[ii +  1] = newValue[0];   // Status
          statusHistory[ii +  2] = formattedDate; // Date
          break;
        }
      }

      range_statusHistory.setNumberFormat('@').setValues([statusHistory]);

      if (statusHistory[28].split(' / ', 1)[0] === 'Complete') // Current Status is Complete
      {
        const completedWarrantiesSheet = spreadsheet.getSheetByName('Completed_Warranties')
        values.push(formattedDate, 'Tag# ' + values[0] + ' - ' + values[1] + ' - ' + formatDate(values[9], spreadsheet) + ' - ' + formattedDate)

        const lastRow = completedWarrantiesSheet.getLastRow();
        const lastCol = completedWarrantiesSheet.getLastColumn() - 1; // The minus one is for the array indexing for the sorting function just below
        const completedOrders = completedWarrantiesSheet.getSheetValues(1, 1, lastRow, lastCol + 1)
        const header = completedOrders.shift()

        const completedOrders_Sorted = Array.from(new Set(completedOrders.concat([values]).sort((a, b) => (a[lastCol] < b[lastCol]) ? 1 : -1).map(JSON.stringify)), JSON.parse)
        completedOrders_Sorted.unshift(header)

        const numCompletedOrders = numFormats.unshift(new Array(lastCol + 1).fill('@'));
        completedWarrantiesSheet.getRange(1, 1, numCompletedOrders, lastCol + 1).setNumberFormat('@').setValues(completedOrders_Sorted)

        const remainingWarranties = allActiveWarranties.filter(tag => tag[0] !== tagNum);
        allActiveWarrantiesSheet.deleteRow(numRows + 1).getRange(2, 1, numRows - 1, remainingWarranties[0].length).setNumberFormat('@').setValues(remainingWarranties)

        const numCols = statusPage.getLastColumn();
        const remainingWarranties_StatusPage = statusPage.getSheetValues(3, 1, numRows, numCols).filter(tag => tag[0] !== tagNum);
        statusPage.deleteRow(numRows + 2).getRange(3, 1, numRows - 1, numCols).setValues(remainingWarranties_StatusPage)
      }
      // else if (statusHistory[28] === 'Sent to Parksville for Repair')
      // {
      //   const previousStatus = values[ii + 24];

      //   if (previousStatus === 'Accepted in Richmond')
      //     addWarrantyToTransferSheet(values, 'Richmond', spreadsheet)
      //   else if (previousStatus === 'Accepted in Rupert')
      //     addWarrantyToTransferSheet(values, 'Rupert', spreadsheet)
      //   else
      //   {
      //     const ui = SpreadsheetApp.getUi()
      //     var response = ui.prompt('Adding Items to the Transfer Sheet', 
      //       'Please type the location name that the items are shipping from initially (Ignore upper or lowercase): \"rich" or \"pr\".', ui.ButtonSet.OK_CANCEL);

      //     var textResponse = response.getResponseText().toUpperCase();

      //     if (textResponse == 'RICH')
      //       addWarrantyToTransferSheet(values, 'Richmond', spreadsheet)
      //     else if (textResponse == 'PR')
      //       addWarrantyToTransferSheet(values, 'Rupert', spreadsheet)
      //     else
      //       ui.alert('Your typed response did not exactly match any of the location choices. Please Try again.')
      //   }
      // }
      // else if (statusHistory[28] === 'Sent Back to Richmond for Distribution to Customer')
      //   addWarrantyToTransferSheet(values, 'Parksville', spreadsheet)
      else if (statusHistory[28] === 'Accepted in Richmond' || statusHistory[28] === 'Accepted in Parksville' || statusHistory[28] === 'Accepted in Rupert')
        allActiveWarrantiesSheet.getRange(row, firstColumnForStatusData).setNumberFormat('@').setValue(values[11]);
      else
        manageDataValidation('Status', newValue[0], spreadsheet); // Add a new status if the user is free typing in the cell

      return;
    case 8: // Address
      col = [4];
      newValue[0] = toProper(newValue[0])
      range.setValue(newValue[0])
      break;
    case 9: // City
      col = [5];
      newValue[0] = toProper(newValue[0])
      range.setValue(newValue[0])
      break;
    case 10: // Province
      col = [6];
      newValue[0] = newValue[0].toString().toUpperCase();
      range.setValue(newValue[0])
      break;
    case 11: // Phone
      col = [8];
      newValue[0] = formatPhoneNumber(newValue[0]);
      range.setValue(newValue[0])
      break;
    case 12: // Email
      col = [9];
      newValue[0] = newValue[0].toString().toLowerCase();
      range.setValue(newValue[0])
      break;
    case 13: // Items
      newValue = newValue[0].split('\n');
      const numItems = newValue.length

      if (numItems < 6)
        newValue.push(...new Array(6 - numItems).fill(''))
      else if (numItems > 6)
      {
        range.setValue('')
        newValue = new Array(6).fill('')
        SpreadsheetApp.getUi().alert('Six items maximum. Please reduce the number of rows.')
      }

      col = newValue.map((_,n) => n + 20);
      break;
    case 14: // Type
      col = [11];
      newValue[0] = toProper(newValue[0]);
      range.setValue(newValue[0])
      break;
  }

  col.map((c, i) => allActiveWarrantiesSheet.getRange(row, c).setNumberFormat('@').setValue(newValue[i]))
}