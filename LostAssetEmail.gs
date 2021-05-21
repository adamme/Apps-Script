// sheet rows: Timestamp	Email Address	First Name	Last Name	Managers Email	Building You Sit At	Current Best Contact Method	Cruise Asset(s) Lost/Stolen	Specifics About Lost/Stolen Items	Location Items Stolen or Last Seen	Lost or Stolen	Badge Lost/Stolen	Badge Only			emailSent		

function sendLostEmail() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Form Responses 1")
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var header = data.shift();

  var email_sent_col = header.indexOf("emailSent");


  var obj = data.map(function (values) {
    return header.reduce(function (o, k, i) {
      o[k] = values[i];
      return o;
    }, {});
  });

  obj.forEach(function (row, rowIdx) {
    var asset = row["Asset(s) Lost/Stolen"];
    var emailTo = row["Email Address"];
    var eamilCCList = row["Managers Email"] + ", " + row["Email Address"]
    var emailBCCList = row["Managers Email"] + ",itsupport@domain.com,security@domain.com,gsoc@domain.com,lost@domain.pagerduty.com";
    var subject = "";
    var lostorstolen = row["Lost or Stolen"];
    var htmlbody = "Reporter: <br>" + row["First Name"] + " " + row["Last Name"] + "<br><br>" +
      "Date: <br>" + row["Timestamp"] + "<br><br>" +
      "Managers Email: <br>" + row["Managers Email"] + "<br><br>" +
      "Building You Sit At: <br>" + row["Building You Sit At"] + "<br><br>" +
      "Current Best Contact Method: <br>" + row["Current Best Contact Method"] + "<br><br>" +
      "Lost or Stolen: <br>" + row["Lost or Stolen"] + "<br><br>" +
      "Asset(s) Lost/Stolen: <br>" + row["Asset(s) Lost/Stolen"] + "<br><br>" +
      "Specifics About Lost/Stolen Items: <br>" + row["Specifics About Lost/Stolen Items"] + "<br><br>" +
      "Location Items Stolen or Last Seen: <br>" + row["Location Items Stolen or Last Seen"]
    if (asset == "Badge" && lostorstolen == "Lost") {
      try {
        subject = "[Lost] asset reported lost/stolen: " + row["First Name"] + " " + row["Last Name"];
        GmailApp.sendEmail("gsoc@domain.com", subject, "", { cc: eamilCCList,  htmlBody: htmlbody, from: "lost@domain.com" })
        data[rowIdx][email_sent_col] = new Date();
      } catch (e) {
        data[rowIdx][email_sent_col] = e.message;
      }
    } else if (asset == "Badge" && lostorstolen == "Stolen") {
      try {
        subject = "[Stolen] asset reported lost/stolen: " + row["First Name"] + " " + row["Last Name"];
        GmailApp.sendEmail("gsoc@domain.com", subject, "", { cc: eamilCCList, htmlBody: htmlbody, from: "lost@domain.com" })
        data[rowIdx][email_sent_col] = new Date();
      } catch (e) {
        data[rowIdx][email_sent_col] = e.message;
      }
    }
    if (row.emailSent === "" && asset != "Badge" && lostorstolen == "Lost") {
      try {
        subject = "[Lost] asset reported lost/stolen: " + row["First Name"] + " " + row["Last Name"];
        GmailApp.sendEmail(emailTo, subject, "", { bcc: emailBCCList, htmlBody: htmlbody, from: "lost@domain.com" })
        data[rowIdx][email_sent_col] = new Date();
      } catch (e) {
        data[rowIdx][email_sent_col] = e.message;
      }
    } else if (row.emailSent === "" && asset != "Badge" && lostorstolen == "Stolen") {
      try {
        subject = "[Stolen] asset reported lost/stolen: " + row["First Name"] + " " + row["Last Name"];
        GmailApp.sendEmail(emailTo, subject, "", { bcc: emailBCCList, htmlBody: htmlbody, from: "lost@domain.com" })
        data[rowIdx][email_sent_col] = new Date();
      } catch (e) {
        data[rowIdx][email_sent_col] = e.message;
      }
    }
  });

  dataRange.offset(1, 0, data.length).setValues(data);
}
