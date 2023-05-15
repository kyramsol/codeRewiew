var fieldsObject = {
  "Name": [14],
  "Accrual rate": [7, 11],
  "Tax": [7, 11],
  "Salary Category for P&L": [23],
  "Tax Category for P&L": [23],
  "Additional Column 1": [18],
  "Additional Column 2": [19],
  "Additional Column 3": [20],
  "Additional Column 4": [21],
  "Additional Column 5": [22]
}


/**
 * Processes payroll data into Database transactions format.
 *    
 * Return {null}
 */


function database_format() {


  // Database sheet preparation

  // 1. Database {sheet} initialization 
  const databaseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");


  // 2. Remove setted {filter} for existed transactions. => Depreciate setValue problem.
  try {
    databaseSheet.getRange("A1:Z").getFilter().remove();
  }
  catch (e) {
    Logger.log(e)
  }


  // Till Date initialization {date} => For end period needed

  //1. Initialize limitation period  => Limition till what month we need to create transactions

  //  By Default, means that we create transactions till the end of the current month. 
  // If you need additional month to get - just change this parameter. 

  // Example: 
  // monthLimitation = 2 (form transaction till the end of the second month from now)

  var monthLimitation = 0

  //2. Create limitation {date} => Limition till what month we need to create transactions

  var limitationDate = new Date(new Date().getFullYear(), new Date().getMonth() + 1 + monthLimitation, 1);


  //  Getting existed transaction from Database {sheet} 
  // - removing Claims and Addional payment transactions
  // - adding row index at the end of each existed transaction 
  var existedTransactions = databaseSheet.getRange("B2:Y").getValues().filter(x => x[0] != "" && String(x[0]).length != 36)
  Logger.log(existedTransactions)

  // Getting {array} of ids from existedTransactions
  var existedIds = existedTransactions.map(x => x[0])


  // Preparing edit logs for processing 

  // 1. Initialize {sheet} with Edit logs   
  var logEditSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Edits Log");

  // 2. Filtering new Edit logs and adding {int} row index per each log
  // - "When change?" value does not exist
  // - remove empty rows
  var logsList = logEditSheet.getRange("A2:F").getValues().map((x, i) => [...x, i + 2]).filter(x => x[5] == "" && x[0] != "")
  Logger.log(logsList)


  // Payments tab preparing and processing

  //// Preparation 

  // 1. Initialize Payment {sheet} 
  var paymentdatabaseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payment');

  // 2. Identifying lastRow of the Payment {sheet} 
  var lastRow = paymentdatabaseSheet.getRange('A4:A').getValues().filter(x => x[0] != "").length


  // 3. Initialize {map} from Handbook {sheet} to identify the transaction nature during processing
  var handbook = new Map(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('General Settings').getRange('G24:H36').getValues());

  var reportCurrency = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('General Settings').getRange('C21').getValue();



  // 4. Getting the row with paymentsMonths and identify limitation using Limitation month {date}
  var paymentsMonths = String(paymentdatabaseSheet.getRange(1, 13, 1, paymentdatabaseSheet.getLastColumn() - 13).getValues()[0].map(x => !x ? "" : Utilities.formatDate(new Date(x), "GMT", "MM-yyyy"))).split(",");
  var limitationIndex = paymentsMonths.indexOf(String(Utilities.formatDate(new Date(limitationDate), "GMT", "MM-yyyy")))

  // 5. Getting the {array} with the employees information 
  var emplInfo = paymentdatabaseSheet.getRange(4, 1, lastRow, 11).getValues();

  // 6. Getting the {array} with the paymentsMonths to set
  var monthList = paymentdatabaseSheet.getRange(4, 13, lastRow, limitationIndex + 13).getValues();
  Logger.log(monthList.length)


  // 7. Getting entry month {range} => For further offsetting
  var beginDate = paymentdatabaseSheet.getRange(1, 13);
  Logger.log(beginDate.offset(0, 0).getValue().getTime())




  //// Processing

  var newTransactions = []

  // 1. Mapping through the patments {sheet} months

  paymentsMonths.map((p, i) => {

    // - using 10  offset iterator - 10 is the number of columns between each month block on the Payments sheet
    // - check only for new Intenal ids !existedIds.includes(`${x[0]}-${i + 13}`)
    // - transactionDate is not more than limitationDate
    if (!i % 10 && new Date(beginDate.offset(0, i).getValue()) < new Date(limitationDate)) {

      // - initialize Date value for the transaction and get Timestamp from it

      var transactionDate = beginDate.offset(0, i).getValue().getTime()

      // 1.1 Nested mapping through Employees of the iterated Date (for loop above)

      emplInfo.map((x, y) => {
        // - identifying of the Category is connected with the Project time => return BS Date index, on opposite - PL Date
        var dateIndex = (!!handbook.get(x[4]) ? 6 : 5);

        // - intiliaze transaction details according to the emplInfo x values
        var salaryAmount = monthList[y][i]
        var taxAmount = monthList[y][i + 3]
        var internalIDSalary = `${x[0]}-${i + 13}`
        var internalIDSalaryTax = `${x[0]}-${i + 16}`

        // - creating transaction settings {object} using conditional pushing of the properties (where we check if there is new transaction and if salary/tax transaction is > 0 )
        var transactionSettings = {
          ...(!existedIds.includes(internalIDSalary) && !!salaryAmount && typeof monthList[y][i] == "number") && { salary: [salaryAmount, internalIDSalary] },
          ...(!existedIds.includes(internalIDSalaryTax) && !!taxAmount && typeof monthList[y][i + 3] == "number") && { tax: [taxAmount, internalIDSalaryTax] },
        }
        Logger.log(transactionSettings)


        // - mapping within transactionSettings keys and create transactions according to the keys parameters 
        Object.keys(transactionSettings).map(l => {

          // - identify the index of what category we need to get Salary or Tax
          var categoryIndex = (l == "salary" ? 4 : 5)

          // - transaction creating using transactionSettings parameters, emplInfo x values
          var transaction = new Array(24).fill("");
          transaction[0] = transactionSettings[l][1] // Payroll internal ID - [EmployeeID]-[MonthRowIndex] {timestamp}
          transaction[1] = Utilities.getUuid() // id {UUID}
          transaction[2] = new Date() // Created {timestamp}
          transaction[parseInt(dateIndex)] = Utilities.formatDate(new Date(transactionDate), "GMT+3", "MM/dd/yyyy") //BS Date or PL Date {date}
          transaction[7] = -transactionSettings[l][0] // Amount {float}
          transaction[8] = "General Payroll" // Account {string}
          transaction[9] = reportCurrency // Report currency {float}
          transaction[10] = 1 // Exchange rate {float}
          transaction[11] = -transactionSettings[l][0] // Amount {float}
          transaction[12] = "Salary for " + Utilities.formatDate(new Date(transactionDate), "GMT+3", "MMM-yy") // Purpose {string}
          transaction[14] = x[2] // Employee {string}
          transaction[18] = x[6] // Extra Data 1 {string}
          transaction[19] = x[7] // Extra Data 2 {string}
          transaction[20] = x[8] // Extra Data 3 {string}
          transaction[21] = x[9] // Extra Data 4 {string}  
          transaction[22] = x[10] // Extra Data 5 {string}
          transaction[23] = x[parseInt(categoryIndex)] // Salary category {string}                      

          newTransactions.push(transaction)
        })
      })
    }
  })



  // Apply new log changes for the existedTransactions {array}

  // 1. forEach iteration within the new log changes
  logsList.forEach(x => {
    // Nested map to get the indexes of transactions we need to change connected with x forEach log change
    var indexesToChange = existedIds.map((e, i) => [(String(e).includes(x[2]) ? i : "")]).filter(f => f[0] != "");

    Logger.log(indexesToChange)

    Logger.log(fieldsObject[x[4]])

    // forEach iteration within the indexesToChange and apply the change x forEach log change
    indexesToChange.forEach((e) => {
      // Using type of change x[4] as a key - we get the value {array} from the fieldsObject {object} with indexes of existedTransaction to change and add the Edited {timestamp}
      fieldsObject[x[4]].forEach(f => existedTransactions[e][parseInt(f)] = x[3])
      existedTransactions[e][3] = new Date()
    })
    SpreadsheetApp.flush()
    // Set the {timestamp} in the Edits Log {sheet} that the change x is applied
    if (!indexesToChange.length) {
      logEditSheet.getRange(parseInt(x[6]), 6).setValue("Not found IDs!")
    }
    else { logEditSheet.getRange(parseInt(x[6]), 6).setValue(new Date()) }

    SpreadsheetApp.flush()
  })

  // Concatanate newTransaction {array} with the existedTransactions (with applies logs) {array}

  var arrayToSet = newTransactions.concat(existedTransactions)


  // Claims and Bonuses Import processing


  // 1. Initialize Claims and Additional payment transactions {array}
  var claims = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claims and Bonuses Import').getRange('B3:R').getValues();

  // 2. Claims and Bonuses transaction formation
  // - filtered empty rows
  // - mapping within the transaction
  // - forming Database transaction and pussing them to the arrayToSet {array}
  claims.filter(x => x[0] != "").map(x => {
    if (new Date(x[3]) < new Date(limitationDate)) {
      var transaction = new Array(24).fill("");
      transaction[0] = x[0]
      transaction[1] = x[0]
      transaction[2] = x[1]
      transaction[3] = x[2]
      transaction[5] = x[3]
      transaction[7] = -x[5]
      transaction[8] = "General Payroll"
      transaction[9] = x[5]
      transaction[10] = x[7] / x[5]
      transaction[11] = -x[7]
      transaction[12] = x[8]
      transaction[16] = x[11]
      transaction[17] = x[10]
      transaction[18] = x[12]
      transaction[19] = x[13]
      transaction[20] = x[14]
      transaction[21] = x[15]
      transaction[22] = x[16]
      transaction[23] = x[9]
      arrayToSet.push(transaction)
    }
  })

  // Sorting arrayToSet by the P&L Date

  arrayToSet = arrayToSet.sort((a, b) => new Date(a[4]).getTime() - new Date(b[4]).getTime())

  // Clear Database {sheet} old content
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database").getRange(2, 2, SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database").getLastRow(), 24).clearContent()

  Logger.log(arrayToSet)

  // Set arrayToSet {array} to the Database {sheet}
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database").getRange(2, 2, arrayToSet.length, arrayToSet[0].length).setValues(arrayToSet)

  // Create filter for the Database {sheet}
  databaseSheet.getRange("A1:Z").createFilter()

}




