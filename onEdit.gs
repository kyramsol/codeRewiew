
/**
 * onEdit simple function that executes:
 * 
 * @param {event} e - event object 
 * - Logging changes of Employees list {sheet} Name, Categories, Additional columns changes
 * - Adding new row for the Employees list {sheet}
 * - Initialize IDs for the unique Employees list {sheet} row => further using for Internal ID generation
 * - Initialize Claims and Bonuses General Admin {sheet} transactions UUID, Create {timestamp}, Change {timestamp}
 * 
 *
 */


function onEdit(e) {
  // Initialize event {object} values
  const activeCell = e.range;
  const value = e.value;
  const oldValue = e.oldValue;
  const activeSheet = activeCell.getSheet()
  const sheetName = activeCell.getSheet().getName();
  const activeRow = activeCell.getRow();
  const activeColumn = activeCell.getColumn();
  const databaseExisted = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database").getRange("B2:B").getValues().filter(x => x[0] != "");
  const logsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Edits Log");


  // Assigning separate functions per sheetName to execute
  const functionHandbook = {

    "Employees list": employeesList,
    "Accrual Admin": adminAccrual,
    "Claims and Bonuses Admin": claimsBonuses,
    "Payment": taxesLog

  }

  // Check for the oldValue is not equal to value to moving on
  if (value != oldValue) {
    // - execute the function from the object, using sheetName as a key
    functionHandbook?.[String(sheetName)](activeSheet, activeRow, activeColumn, databaseExisted, value, logsSheet)

  }

}



/**
 * Function for the Employees List {sheet} that executes: 
 * 
 * - Logging changes of Name, Categories, Additional columns changes;
 * - Adding new row if needed;
 * - Initialize internal IDs for the unique row => further using for Internal ID generation
 * 
 * @param {sheet} activeSheet - where script executes
 * @param {int} activeRow - active row number (! not index !) from the event object
 * @param {int} activeColumn - active column number (! not index !) from the event object
 * @param {int} databaseExisted - database existed transactions
 * @param {string} || {int}  value - event object value 
 * @param {sheet} logsSheet - edit logs storage
 * 
 * Return {null}
 * 
 */




function employeesList(activeSheet, activeRow, activeColumn, databaseExisted, value, logsSheet) {

  var headerName = activeSheet.getRange(2, activeColumn).getValue();
  var id = activeSheet.getRange(activeRow, 1).getValue();
  var headersToCheck = ["Name", "Salary Category for P&L", "Tax Category for P&L", "Additional Column 1", "Additional Column 2", "Additional Column 2", "Additional Column 3", "Additional Column 4", "Additional Column 5"]


  if (activeRow > 2 && activeColumn > 2) {

    Logger.log(`Active row >2, column >2`)


    /// 1. Checking for the Employees list {sheet} value change case
    // - header of the active cell includes into  checkToHeader {array}
    // - databaseExisted {array} contains the transaction with the Employees list {sheet} row ID where change was made
    if (headersToCheck.includes(headerName) && !!databaseExisted.filter(x => String(x[0]).includes(id)).length && activeSheet.getRange(activeRow, 1).getValue()) {

      Logger.log(`Active cell in Headers to check`)
      var record = [[new Date(), "Current", id, value, headersToCheck.filter(x => x == headerName)[0]]]
      logsSheet.getRange(logsSheet.getLastRow() + 1, 1, record.length, record[0].length).setValues(record)
      SpreadsheetApp.flush()
    }


    /// 2. Checking fot the Employees list {sheet} new row entering case => Generate row ID
    //  - expected cell for the Employee name entering is not equal to the default value "Add Name";
    //  - row ID for the active row did not created before 

    else if (activeColumn == 3 && activeRow > 2 && activeSheet.getRange(activeRow, 3).getValue() != "Add Name" && !activeSheet.getRange(activeRow, 1).getValue()) {
      Logger.log(`ID initialize`)
      SpreadsheetApp.flush()
      activeSheet.getRange(activeRow, 1).setValue(generateString(6))
      SpreadsheetApp.flush()
    }

    /// 3. Checking for the Employees list {sheet} new row appending case 
    else if (activeSheet.getLastRow() == activeRow && activeSheet.getRange(activeRow, 3).getValue() !== "Add name") {
          Logger.log(`New row appending`)

      // - create an empty array for the appended row and set default values for this row
      let row = new Array(23).fill("")
      row[2] = "Add name"
      row[3] = "Add position"
      row[5] = "Add direction"
      row[6] = "Add category"
      row[7] = "Add category"
      row[8] = "Add Start date"
      row[9] = "Still running"
      row[10] = "Add Type"
      row[12] = "Add Office"
      row[14] = "Add Email"
      row[15] = "Add Comment"

      activeSheet.appendRow(row);
      activeSheet.getRange('A1:V').getFilter().remove().createFilter();
      SpreadsheetApp.flush()

    }

  }

}



/**
 * Function for the Accrual Admin {sheet} that executes: 
 * 
 * - Logging changes of Accrual rate;
 * 
 * @param {sheet} activeSheet - where the script executes
 * @param {int} activeRow - active row number (! not index !) from the event object
 * @param {int} activeColumn - active column number (! not index !) from the event object
 * @param {int} databaseExisted - database existed transactions
 * @param {string} || {int}  value - event object value 
 * @param {sheet} logsSheet - edit logs storage
 * 
 * Return {null}
 * 
 */

function adminAccrual(activeSheet, activeRow, activeColumn, databaseExisted, value, logsSheet) {

  // Initiliaze the row ID of the Employees list {sheet} reference
  const id = activeSheet.getRange(activeRow, 1).getValue();

  // Getting headers with dates from the Payments {sheet} and creating MM-yyyy format {array}
  // - getting only {date} values for mutation
  const dateList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payment").getRange("A1:IQ1").getValues()[0].map(x => (x instanceof Date) ? `${new Date(x).getMonth() + 1}-${new Date(x).getFullYear()}` : "");

  // Getting the active row header from the Accrual Admin {sheet} and create the MM-yyyy {string}
  const date = `${new Date(activeSheet.getRange(1, activeColumn).getValue()).getMonth() + 1}-${new Date(activeSheet.getRange(1, activeColumn).getValue()).getFullYear()}`;

  // Identifying the index of the date {str} in the Payment {sheet}
  const dateIndex = String(dateList).split(",").indexOf(date);

  // Checking for the editing Accrual Rate value
  // - expected ID of this transaction is includes in the databaseExisted {array}
  if (activeColumn > 2 && activeRow > 2 && activeSheet.getRange(2, activeColumn).getValue() == "Accrual Rate" && !!databaseExisted.filter(x => String(x[0]).includes(`${id}-${dateIndex + 1}`)).length && !!dateIndex) {

    // Initialize a record with log change info
    const record = [[new Date(), "Current", `${id}-${dateIndex + 1}`, value, "Accrual rate"]]
    // Append the record {array} to the Edits Log {sheet}
    logsSheet.getRange(logsSheet.getLastRow() + 1, 1, record.length, record[0].length).setValues(record)
  }
}


/**
 * Function for the Claims and Bonuses Admin {sheet} that executes: 
 * 
 * - Inintialize UUID for the new transaction;
 * - Tag  Created {timestamp} and Edited {timestamp} parameters 
 * 
 * @param {sheet} activeSheet - where script executes
 * @param {int} activeRow - active row number (! not index !) from the event object
 * @param {int} activeColumn - active column number (! not index !) from the event object
 * 
 * Return {null}
 * 
 */


function claimsBonuses(activeSheet, activeRow, activeColumn, databaseExisted, value, logsSheet) {


  // Inirialize the row index to check {array}
  const rowsTocheck = [5, 6, 7, 8, 11];

  // Checking for rows from rowsTocheck {array} are not empty case
  // Return {boolean}
  const newCheck = rowsTocheck.every(x => activeSheet.getRange(activeRow, x).getValue() != "")
  Logger.log(`newCheck:${newCheck}`)
  const id = activeSheet.getRange(activeRow, 2).getValue();
  const created = activeSheet.getRange(activeRow, 3).getValue();

  // Checking that the active value was edited outside the headers and rows items 
  if (activeRow > 2 && activeColumn > 4) {

    // Checking for the new transaction case
    // - id and Created filed are empty
    // - newCheck {boolean} - minimal required fields are not empty
    if (!id + !created == 2 && newCheck) {

      // Setting UUID and created {timestamp} for the newly created transaction 
      activeSheet.getRange(activeRow, 2).setValue(Utilities.getUuid());
      activeSheet.getRange(activeRow, 3).setValue(new Date());

    }

    // Checking for the transaction editing
    else if (!id + !created == 0) {

      // Setting edited {timestamp} value of the transaction
      activeSheet.getRange(activeRow, 4).setValue(new Date());
    }
  }
}

/**
 * Function for the Payments {sheet} that executes: 
 * 
 * - Logging changes of Taxes;
 * 
 * @param {sheet} activeSheet - where script executes
 * @param {int} activeRow - active row number (! not index !) from the event object
 * @param {int} activeColumn - active column number (! not index !) from the event object
 * @param {int} databaseExisted - database existed transactions
 * @param {string} || {int}  value - event object value 
 * @param {sheet} logsSheet - edit logs storage
 * 
 * Return {null}
 * 
 */


function taxesLog(activeSheet, activeRow, activeColumn, databaseExisted, value, logsSheet) {

  const id = activeSheet.getRange(activeRow, 1).getValue();

  // Checking for the case when the value under the Tax header was edited in the Payment {sheet}
  // - header of the active cell is equal to "Tax"
  // - expected transaction ID was existed in the databaseExisted {array}
  if (activeColumn > 12 && activeRow > 3 && activeSheet.getRange(2, activeColumn).getValue() == "Tax" && !!databaseExisted.filter(x => String(x[0]).includes(`${id}-${activeColumn}`)).length) {

    // Initialize a record with log change info
    const record = [[new Date(), "Current", `${id}-${activeColumn}`, value, "Tax"]]

    // Append the record {array} to the Edits Log {sheet}
    logsSheet.getRange(logsSheet.getLastRow() + 1, 1, record.length, record[0].length).setValues(record)
  }
}



/**
 * Function for the Unique Employee List {sheet} row id generating 
 * 
 * @param {int} length - length of an ID needed 
 * 
 * Return {string} ID with the length set in the input that contains digits, lower/upper case letters
 * 
 */

function generateString(length) {

  // Initialize the {string} with all available characters for the ID
  const characters = 'ABCDEFGHIJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  const charactersLength = characters.length;

  let result = new Array(length).fill(1).reduce(acc => (acc += characters.charAt(Math.floor(Math.random() * charactersLength))))
  Logger.log(result)
  return result;
}




