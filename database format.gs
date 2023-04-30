
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



function database_format() {

  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");
  try { ss.getRange("A1:Z").createFilter() }
  catch { }
  var filter = ss.getFilter();
  filter.remove();

  var month_fortoday = new Date(new Date().getFullYear(), new Date().getMonth() + 1, 1);
  var current_month = month_fortoday.getMonth();
  var current_year = month_fortoday.getFullYear();

  var existedTransactions = ss.getRange("B2:Y").getValues().filter(x => x[0] != "" && ["Bonus", "Additional payment"].includes(x[12]) == false).map((x, i) => [...x, i])
  Logger.log(existedTransactions)
  var existedIds = existedTransactions.map(x => x[0])


  var logsSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Edits Log")
  var logsList = logsSS.getRange("A2:F").getValues().map((x, i) => [...x, i + 2]).filter(x => x[5] == "" && x[0] != "")
  Logger.log(logsList)



  var paymentSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payment');
  var lr = paymentSS.getRange('A4:A').getValues().filter(x => x[0] != "").length
  var handbook = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('General Settings').getRange('G24:G36').getValues();
  var handbookcheck = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('General Settings').getRange('H24:H36').getValues();
  var reportCurrency = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('General Settings').getRange('C21').getValue();

  var months = String(paymentSS.getRange(1, 13, 1, paymentSS.getLastColumn() - 13).getValues()[0].map(x => !x ? "" : `${String(new Date(x).getMonth())}.${String(new Date(x).getFullYear())}`)).split(",");
  var index_source = months.indexOf(`${String(current_month)}.${String(current_year)}`)

  var emplInfo = paymentSS.getRange(4, 1, lr, 6).getValues();
  var monthList = paymentSS.getRange(4, 13, lr, index_source + 13).getValues();
  Logger.log(monthList.length)




  var beginDate = paymentSS.getRange(1, 13);
  Logger.log(beginDate.offset(0, 0).getValue().getTime())


  var array = []

  // Payment Tab Import

  for (i = 0; monthList[0].length > i; i += 10) {



    if (new Date(beginDate.offset(0, i).getValue()) < new Date(current_year, current_month, 1)) {


      Logger.log(i)



      var dateval = beginDate.offset(0, i).getValue().getTime()

      emplInfo.map((x, y) => {


        var index = String(handbook).split(",").indexOf(x[4]);
        if (handbookcheck[index][0] == true) {
          if (!monthList
          [y][i] == false && typeof monthList
          [y][i] == "number") {
            array.push(
              [`${x[0]}-${i + 13}`,
                "",
                "",
                "",
                "",
                "",
              Utilities.formatDate(new Date(dateval), "GMT+3", "MM/dd/yyyy"),
              -monthList
              [y][i],
                "General Payroll",
                reportCurrency,
                "1",
              -monthList
              [y][i],
              "Salary for " + Utilities.formatDate(new Date(dateval), "GMT+3", "MMM-yy"),
                "",
              x[2],
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
              x[4],
              ])
          };

          if (!monthList
          [y][i + 3] == false && typeof !monthList
          [y][i + 3] == "number") {
            array.push(
              [`${x[0]}-${i + 13 + 3}`,
                "",
                "",
                "",
                "",
                "",
              Utilities.formatDate(new Date(dateval), "GMT+3", "MM/dd/yyyy"),
              -monthList
              [y][i + 3],
                "General Payroll",
                reportCurrency,
                "1",
              -monthList
              [y][i + 3],
              "Salary tax for " + Utilities.formatDate(new Date(dateval), "GMT+3", "MMM-yy"),
                "",
              x[2],
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
              x[5],
              ])
          }
        }

        else {
          if (!monthList
          [y][i] == false && typeof monthList
          [y][i] == "number") {
            array.push([
              `${x[0]}-${i + 13}`,
              "",
              "",
              "",
              "",
              Utilities.formatDate(new Date(dateval), "GMT+3", "MM/dd/yyyy"),
              "",
              -monthList
              [y][i],
              "General Payroll",
              reportCurrency,
              "1",
              -monthList
              [y][i],
              "Salary for " + Utilities.formatDate(new Date(dateval), "GMT+3", "MMM-yy"),
              "",
              x[2],
              "",
              "",
              "",
              "",
              "",
              "",
              "",
              "",
              x[4],
            ])
          }

          if (!monthList
          [y][i + 3] == false && typeof monthList
          [y][i + 3] == "number") {
            array.push(
              [
                `${x[0]}-${i + 13 + 3}`,
                "",
                "",
                "",
                "",
                Utilities.formatDate(new Date(dateval), "GMT+3", "MM/dd/yyyy"),
                "",
                -monthList
                [y][i + 3],
                "General Payroll",
                reportCurrency,
                "1",
                -monthList
                [y][i + 3],
                "Salary tax for " + Utilities.formatDate(new Date(dateval), "GMT+3", "MMM-yy"),
                "",
                x[2],
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                x[5],
              ])
          }
        };
      })
    }
  }

  var newTransactions = array.filter(x => !existedIds.includes(x[0]))
  newTransactions.forEach((x) => {
    x[1] = Utilities.getUuid();
    x[2] = new Date()
  })

  Logger.log(newTransactions)

  // Apply log changes

  logsList.forEach(x => {
    let indexesTochange = existedTransactions.flatMap((e, i) => String(e[0]).includes(x[2]) ? i : []);

    Logger.log(fieldsObject[x[4]])
    indexesTochange.forEach((e) => {
      fieldsObject[x[4]].forEach(f => existedTransactions[e][parseInt(f)] = x[3])
      existedTransactions[e][3] = new Date()
    })
    SpreadsheetApp.flush()
    logsSS.getRange(parseInt(x[6]), 6).setValue(new Date())
    SpreadsheetApp.flush()
  })
  existedTransactions = existedTransactions.map(x => x.slice(0, -1))

  array = newTransactions.concat(existedTransactions)


  // Claims and Bonuses Import

  var claims = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claims and Bonuses Import').getRange('B3:R').getValues();

  claims.filter(x => x[0] != "").map(x => {
    if (new Date(x[3]) <= new Date(current_year, current_month, 0)) { array.push([x[0], x[0], x[1], x[2], "", x[3], "", -x[5], "General Payroll", x[5], x[7] / x[5], -x[7], x[8], "", "", "", x[11], x[10], x[12], x[13], x[14], x[15], x[16], x[9]]) }
  })

  array = array.sort((a, b) => new Date(a[4]).getTime() - new Date(b[4]).getTime())

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database").getRange(2, 2, SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database").getLastRow(), SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database").getLastColumn()).clearContent()

  Logger.log(array)

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database").getRange(2, 2, array.length, array[0].length).setValues(array)

  ss.getRange("A1:Z").createFilter()

}


const test = () => {

  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");

  var existedTransactions = ss.getRange("B2:Y").getValues().filter(x => x[0] != "" && ["Bonus", "Additional payment"].includes(x[12]) == false).map((x, i) => [...x, i])
  Logger.log(existedTransactions)

  var logsSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Edits Log")
  var logsList = logsSS.getRange("A2:F").getValues().map((x, i) => [...x, i + 2]).filter(x => x[5] == "" && x[0] != "")
  Logger.log(logsList)

  logsList.forEach(x => {
    let indexesTochange = existedTransactions.flatMap((e, i) => String(e[0]).includes(x[2]) ? i : []);

    Logger.log(fieldsObject[x[4]])
    indexesTochange.forEach((e) => {
      fieldsObject[x[4]].forEach(f => existedTransactions[e][parseInt(f)] = x[3])
      existedTransactions[e][3] = new Date()
    })
    SpreadsheetApp.flush()
    Logger.log(parseInt(x[6]))
    logsSS.getRange(parseInt(x[6]), 6).setValue(new Date())
    SpreadsheetApp.flush()
  })
  Logger.log(existedTransactions)


}



