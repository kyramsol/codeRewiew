function to_key_pair(keys, commonKeys) {

  var result = {}
  var currentKey
  var currentVal

  for (i = 0; i < keys.length; i++) {
    currentKey = commonKeys[i];
    currentVal = keys[i];
    result[currentKey] = currentVal;
  }

  return result
}



function send_notifications_2payments() {

  const payment_sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payment');
  const notificationset_sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notification Settings');
  const selected_month = notificationset_sh.getRange("B1").getValue();
  const selected_method = notificationset_sh.getRange("D1").getValue();
  if (selected_method == 'Email') {
    var notificationem_templ = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notification_E-mail structure');
  }
  else if (selected_method == 'Slack') {
    var notificationem_templ = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notification_Slack structure');
  }
  const months_range = String(notificationset_sh.getRange(2, 1, 1, notificationset_sh.getLastColumn()).getValues()).split(",");
  const month_index = months_range.indexOf(selected_month)
  const month_value = notificationset_sh.getRange(4, month_index + 1).getValue()
  Logger.log(month_index)
  const closed_month = notificationset_sh.getRange(3, month_index + 1).getValue()
  Logger.log(closed_month)

  const ready_key = notificationset_sh.getRange(6, 1, notificationset_sh.getLastRow(), notificationset_sh.getLastColumn()).getValues().map(x => [x[2], x[month_index]]).filter(x => x[0] != "");

  Logger.log(ready_key)

  const adress_key = to_key_pair(notificationset_sh.getRange(6, 4, notificationset_sh.getLastRow(), 1).getValues(), notificationset_sh.getRange(6, 3, notificationset_sh.getLastRow(), 1).getValues());

  // Logger.log(ready_key)
  // Logger.log(adress_key)

  // Payments \\

  const payment_month_range = String(payment_sh.getRange("A1:PY1").getValues()).split(",")
  Logger.log(payment_month_range)
  Logger.log(month_value)
  const payment_month_index = payment_month_range.indexOf(month_value.toString())
  Logger.log(payment_month_index)

  const claims_paym = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claims and Bonuses Import').getRange('B3:I').getValues().filter(x => x[0] != '');

  var payments_array = payment_sh.getRange("A4:PY").getValues().filter(x => x[0] != "")
  payments_array = payments_array.map(x => [x[0], x[payment_month_index], x[payment_month_index + 3], x[payment_month_index + 4]])
  Logger.log(payments_array)

  var arr = [];
  payments_array = payments_array.map(ob => {


    if (String(arr.map(x => x[0])).split(",").indexOf(ob[0]) == -1) {

      // Logger.log(String(arr.map(x => x[0])).split(",").indexOf(ob[0]))
      // Logger.log(ob[0])
      // Logger.log(String(arr.map(x => x[0])).split(","))

      var num1 = 0
      var num2 = 0
      var num3 = 0

      payments_array.map(x => {
        if (x[0] == ob[0]) {
          num1 += x[1]
          num2 += x[2]
          num3 += x[3]
        }


        return num1, num2, num3
      })

      arr.push([ob[0], num1, num2, num3])
    }
    return arr
  })

  Logger.log(payments_array)

  if (closed_month != true) {
    if (selected_method == 'Email') {
      const subject_templ = notificationem_templ.getRange('C5').getValue();
      const header1_templ = notificationem_templ.getRange('C9').getValue();
      const header2_templ = notificationem_templ.getRange('C10').getValue();
      const total_amount_templ = notificationem_templ.getRange('C11').getValue();
      const rate_templ = notificationem_templ.getRange('C12').getValue();
      const transaction_det = notificationem_templ.getRange('C13').getValue();
      const total_income_tax = notificationem_templ.getRange('C14').getValue();
      const footer1_templ = notificationem_templ.getRange('C15').getValue();
      const footer2_templ = notificationem_templ.getRange('C16').getValue();
      const cc_adr = String(notificationem_templ.getRange('C6').getValue()).split(",").filter(x => validateEmail(x) == true)

      for (x in ready_key) {

        if (ready_key[x][1] == "Ready to Send") {

          var empl_index = String(payments_array[0].map(x => x[0])).split(",").indexOf(ready_key[x][0]);
          Logger.log(String(payments_array[0].map(x => x[0])).split(","))
          Logger.log(ready_key[x][0])
          Logger.log(`Employees index ${empl_index}`)
          var repl_name = ready_key[x][0];
          var repl_month = Utilities.formatDate(month_value, "GMT", "MMM-yy");
          var repl_total = payments_array[0][empl_index][3].toFixed(2);
          var repl_rate = payments_array[0][empl_index][1].toFixed(2);
          var repl_curr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("General Settings").getRange("C21").getValue();
          var repl_tax = payments_array[0][empl_index][2].toFixed(2)

          var message_keys = {

            _Name_: repl_name,
            _Month_: repl_month,
            _Total_payment_amount_: repl_total,
            _Rate_Amount_: repl_rate,
            _Reporting_Currency_: repl_curr,
            _Income_Tax_Amount_: repl_tax

          }

          Logger.log(message_keys)

          if (transaction_det == 'Bonuses') {

            var for_table = claims_paym.filter(element => { return element[1].split(":")[0] === ready_key[x][0] && element[0].getMonth() === month_value.getMonth() && element[0].getFullYear() === month_value.getFullYear() && element[5] === 'Bonus' }).map(o => [Utilities.formatDate(o[0], "GMT", "yyyy-MM-dd"), o[5], o[7], o[4].toFixed(2)]).sort((a, b) => a[0] - b[0])
          }
          if (transaction_det == 'Additional Payments') {

            var for_table = claims_paym.filter(element => { return element[1].split(":")[0] === ready_key[x][0] && element[0].getMonth() === month_value.getMonth() && element[0].getFullYear() === month_value.getFullYear() && element[5] === 'Additional payment' }).map(o => [Utilities.formatDate(o[0], "GMT", "yyyy-MM-dd"), o[5], o[7], o[4].toFixed(2)]).sort((a, b) => a[0] - b[0])
          }
          if (transaction_det == 'Bonuses/Additional Payments') {
            var for_table = claims_paym.filter(element => { return element[1].split(":")[0] === ready_key[x][0] && element[0].getMonth() === month_value.getMonth() && element[0].getFullYear() === month_value.getFullYear() && (element[5] === 'Bonus' || element[5] === 'Additional payment') }).map(o => [Utilities.formatDate(o[0], "GMT", "yyyy-MM-dd"), o[5], o[7], o[4].toFixed(2)]).sort((a, b) => a[0] - b[0])
          }
          else {
            var for_table = null
          }

          Logger.log(for_table)

          var general_templ = create_html_email(for_table, repl_curr).getContent();
          general_templ = general_templ.replace('{Header 1}', header1_templ).replace('{Header 2}', header2_templ).replace('{Total Amount}', total_amount_templ).replace('{Rate}', rate_templ).replace('{Total Income Tax}', total_income_tax).replace('{Footer 1}', footer1_templ).replace('{Footer 2}', footer2_templ)
          var subject = replaceAll(subject_templ, message_keys)
          general_templ = replaceAll(general_templ, message_keys)

          Logger.log(adress_key[ready_key[x][0]])
          Logger.log(subject)
          Logger.log(general_templ)
          Logger.log(`For setValue ${x + 6} column, ${month_index + 1} row`)
          if (validateEmail(adress_key[ready_key[x][0]].toString()) == true) {
            sendmessage(adress_key[ready_key[x][0]].toString(), cc_adr.toString(), subject, general_templ)
            notificationset_sh.getRange(String(notificationset_sh.getRange(1, 3, notificationset_sh.getLastRow(), 1).getValues()).split(",").indexOf(repl_name) + 1, month_index + 1).setValue("Notification Sent")

          }
          else {
            notificationset_sh.getRange(String(notificationset_sh.getRange(1, 3, notificationset_sh.getLastRow(), 1).getValues()).split(",").indexOf(repl_name) + 1, month_index + 1).setValue("Invalid Email")
          }
        }
      }
    }

    if (selected_method == 'Slack') {
      var token = notificationem_templ.getRange('C4').getValue();
      Logger.log(`Here is token ${token}`)
      for (x in ready_key) {

        const header1_templ = notificationem_templ.getRange('C9').getValue();
        const header2_templ = notificationem_templ.getRange('C10').getValue();
        const total_amount_templ = notificationem_templ.getRange('C11').getValue();
        const rate_templ = notificationem_templ.getRange('C12').getValue();
        const transaction_det = notificationem_templ.getRange('C13').getValue();
        const total_income_tax = notificationem_templ.getRange('C14').getValue();
        const footer1_templ = notificationem_templ.getRange('C15').getValue();
        const footer2_templ = notificationem_templ.getRange('C16').getValue();

        if (ready_key[x][1] == "Ready to Send") {

          var empl_index = String(payments_array[0].map(x => x[0])).split(",").indexOf(ready_key[x][0]);
          Logger.log(String(payments_array[0].map(x => x[0])).split(","))
          Logger.log(ready_key[x][0])
          Logger.log(`Employees index ${empl_index}`)
          var repl_name = ready_key[x][0];
          var repl_month = Utilities.formatDate(month_value, "GMT", "MMM-yy");
          var repl_total = parseFloat(payments_array[0][empl_index][3]).toFixed(2);
          var repl_rate = parseFloat(payments_array[0][empl_index][1]).toFixed(2);
          var repl_curr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("General Settings").getRange("C21").getValue();
          var repl_tax = parseFloat(payments_array[0][empl_index][2]).toFixed(2)

          var message_keys = {

            _Name_: repl_name,
            _Month_: repl_month,
            _Total_payment_amount_: repl_total,
            _Rate_Amount_: repl_rate,
            _Reporting_Currency_: repl_curr,
            _Income_Tax_Amount_: repl_tax

          }

          Logger.log(message_keys)

          if (transaction_det == 'Bonuses') {

            var for_table = claims_paym.filter(element => { return element[1].split(":")[0] === ready_key[x][0] && element[0].getMonth() === month_value.getMonth() && element[0].getFullYear() === month_value.getFullYear() && element[5] === 'Bonus' }).map(o => [Utilities.formatDate(o[0], "GMT", "yyyy-MM-dd"), o[5], o[7], o[4].toFixed(2)]).sort((a, b) => a[0] - b[0])
          }
          if (transaction_det == 'Additional Payments') {

            var for_table = claims_paym.filter(element => { return element[1].split(":")[0] === ready_key[x][0] && element[0].getMonth() === month_value.getMonth() && element[0].getFullYear() === month_value.getFullYear() && element[5] === 'Additional payment' }).map(o => [Utilities.formatDate(o[0], "GMT", "yyyy-MM-dd"), o[5], o[7], o[4].toFixed(2)]).sort((a, b) => a[0] - b[0])
          }
          if (transaction_det == 'Bonuses/Additional Payments') {
            var for_table = claims_paym.filter(element => { return element[1].split(":")[0] === ready_key[x][0] && element[0].getMonth() === month_value.getMonth() && element[0].getFullYear() === month_value.getFullYear() && (element[5] === 'Bonus' || element[5] === 'Additional payment') }).map(o => [Utilities.formatDate(o[0], "GMT", "yyyy-MM-dd"), o[5], o[7], o[4].toFixed(2)]).sort((a, b) => a[0] - b[0])
          }
          else {
            var for_table = null
          }

          Logger.log(for_table)


          var header1 = replaceAll(header1_templ, message_keys);
          var header2 = replaceAll(header2_templ, message_keys);
          var total_amount = replaceAll(total_amount_templ, message_keys);
          var rate = replaceAll(rate_templ, message_keys);
          var total_income = replaceAll(total_income_tax, message_keys);
          var footer1 = replaceAll(footer1_templ, message_keys);
          var footer2 = replaceAll(footer2_templ, message_keys)



          var payload = {

            "channel": adress_key[ready_key[x][0]].toString(),
            "blocks": [
              {
                "type": "header",
                "text": {
                  "type": "plain_text",
                  "text": header1,
                  "emoji": true
                }
              },
              {
                "type": "section",
                "text": {
                  "type": "plain_text",
                  "text": header2,
                  "emoji": true
                }
              },
              {
                "type": "section",
                "text": {
                  "type": "plain_text",
                  "text": total_amount,
                  "emoji": true
                }
              },
              {
                "type": "section",
                "text": {
                  "type": "plain_text",
                  "text": rate,
                  "emoji": true
                }
              },
              {
                "type": "section",
                "text": {
                  "type": "plain_text",
                  "text": total_income,
                  "emoji": true
                }
              },
              for_table == null ? "" :
                {
                  "type": "section",
                  "fields": makeTableSlack2(for_table, repl_curr)

                },
              {
                "type": "section",
                "text": {
                  "type": "plain_text",
                  "text": footer1,
                  "emoji": true
                }
              },
              {
                "type": "section",
                "text": {
                  "type": "plain_text",
                  "text": footer2,
                  "emoji": true
                }
              },
            ]
          }
          sendAlert(payload, token)
          notificationset_sh.getRange(String(notificationset_sh.getRange(1, 3, notificationset_sh.getLastRow(), 1).getValues()).split(",").indexOf(repl_name) + 1, month_index + 1).setValue("Notification Sent")

        }

      }
    }
  }
}



function sendAlert(payload, token) {
  // const webhook = ""; //Paste your webhook URL here
  var options = {
    "method": "post",
    "contentType": "application/json",
    "muteHttpExceptions": true,
    "payload": JSON.stringify(payload),
    headers: { Authorization: `Bearer ${token}` },
  };

  try {
    UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
  }
  catch (e) {
    Logger.log(e)
  }
}






function create_html_email(table_content, currency) {


  return HtmlService.createHtmlOutput().setContent(`
<html> 
<body> 
<h2> <style type="text/css"><!--td {border: 1px solid #cccccc;}br {mso-data-placement:same-cell;}--> </style> <span style="font-family:Arial,Helvetica,sans-serif;"><strong>{Header 1}</strong></span></h2> 
     <p> <style type="text/css"><!--td {border: 1px solid #cccccc;}br {mso-data-placement:same-cell;}--> </style> <strong><span style="font-family:Arial,Helvetica,sans-serif;">{Header 2}</span></strong></p> 
     <p>&nbsp;</p> <p> <style type="text/css"><!--td {border: 1px solid #cccccc;}br {mso-data-placement:same-cell;}--> </style> <span style="font-family:Arial,Helvetica,sans-serif;">{Total Amount}</span></p>
      <p id="abcd"> <style type="text/css"><!--td {border: 1px solid #cccccc;}br {mso-data-placement:same-cell;}--> </style> <span style="font-family:Arial,Helvetica,sans-serif">{Rate}</span></p>
      ${(table_content.length > 0 ? makeTableHTML(table_content, currency) : "")} 
      <p>&nbsp;</p> <p> <style type="text/css"><!--td {border: 1px solid #cccccc;}br {mso-data-placement:same-cell;}-->
</style> <span style="font-family:Arial,Helvetica,sans-serif;">{Total Income Tax}</span></p> 
<p><span style="font-family:Arial,Helvetica,sans-serif;">{Footer 1}</span></p> 
<p><span style="font-family:Arial,Helvetica,sans-serif;">{Footer 2}</span></p> 
</body> 
</html>`)

}




function sendmessage(receiver, ccadr, subject, text_html) {

  MailApp.sendEmail({
    to: receiver,
    cc: ccadr,
    subject: subject,
    htmlBody: text_html
  });
}


function replaceAll(str, mapObj) {
  var re = new RegExp(Object.keys(mapObj).join("|"), "gi");

  return str.replace(re, function (matched) {
    return mapObj[matched];
  });
}

function makeTableHTML(myArray, currency) {

  var total_sum = myArray.map(x => [Number(x[3]).toFixed(2)]).reduce((partialSum, a) => partialSum + Number(a), 0)
  Logger.log(total_sum)

  var result =
    `<p><span style="font-family:Arial,Helvetica,sans-serif">Below you can see the list of Claims and Bonuses:</span></p> <table border="1" cellpadding="1" cellspacing="1"> <thead> <tr> <th scope="col">Date</th>  <th scope="col">Type</th> <th scope="col">Comment</th> <th scope="col">Amount,${currency}</th> 		</tr> </thead>`;
  for (var i = 0; i < myArray.length; i++) {
    result += '<tr style="text-align:center">';
    for (var j = 0; j < myArray[i].length; j++) {
      result += '<td>' + myArray[i][j] + '</td>';
    }
    result += "</tr>";
  }
  result += `<td style="text-align:center"><strong>Total Amount</strong></td> <td style="text-align:center">&nbsp;</td> <td style="text-align:center">&nbsp;</td> <td style="text-align:center"><strong>${myArray.length > 1 ? parseFloat(total_sum).toFixed(2) : myArray[0][3]}</strong></td> </table>`;
  Logger.log(result)
  return result;
}



function makeTableSlack2(myArray, currency) {


  var total_sum = myArray.map(x => [Number(x[3]).toFixed(2)]).reduce((partialSum, a) => partialSum + Number(a), 0.00)
  var obj = [{
    type: "mrkdwn",
    text: "*Date, Comment*"
  },
  {
    type: "mrkdwn",
    text: `*Amount, ${currency}*`
  }];


  for (var i = 0; i < myArray.length; i++) {
    var add_obj = {
      type: "mrkdwn",
      text: `${myArray[i][0]} - ${myArray[i][2]}`
    }

    obj.push(add_obj)

    var add_obj = {
      type: "mrkdwn",
      text: `${myArray[i][3]}`
    }
    obj.push(add_obj)


  }

  obj.push({
    type: "mrkdwn",
    text: "*Total Amount*"
  },
    {
      type: "mrkdwn",
      text: `${total_sum}`
    })

  return obj;
}

function validateEmail(email) {
  var re = /\S+@\S+\.\S+/;
  return re.test(email);
}
