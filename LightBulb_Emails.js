function myFunction() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var emailSentCount = 0;

  // Send Emails for each Sheet
  for (var n=0; n <= 1; n++) {
    
    var activeSheet = sheets[n];
    var lr = activeSheet.getLastRow();

    for (var i=2; i<=lr; i++) {

    var active = activeSheet.getRange(i,1).getValue();
    var studentFN = activeSheet.getRange(i,3).getValue();
    var studentLN = activeSheet.getRange(i,4).getValue();
    var parentFN = activeSheet.getRange(i,5).getValue();
    var parentLN = activeSheet.getRange(i,6).getValue();
    var parentEmail = activeSheet.getRange(i,7).getValue();
    var secondEmail = activeSheet.getRange(i,8).getValue();
    var classDay = activeSheet.getRange(i,9).getDisplayValue();
    var timeZone = activeSheet.getRange(i,10).getDisplayValue();
    var classTime = activeSheet.getRange(i,11).getDisplayValue();
    var meetingLink = activeSheet.getRange(i,12).getValue();
    var worksheetLink = activeSheet.getRange(i,13).getValue();
    var dateEmailSent = activeSheet.getRange(i,2).getValue();
    
    var mailSubject = studentFN + ": " + activeSheet.getName() + " Worksheet + Homework from Lightbulb Tutoring"

    var greeting = "";
    if ((parentFN.length > 0) || (parentLN.length > 0)) {
        greeting = "Hello" + " " + parentFN.toString().replace(" ","") + " " + parentLN.toString().replace(" ","") + ",";
    } 

    var classTimeBasedOnTimeZone = "";
    if (classTime == "2-3 PM") {
      if (timeZone == "P") {
          classTimeBasedOnTimeZone = "12-1 PM";
      } else if (timeZone == "E") {
          classTimeBasedOnTimeZone = "3-4 PM";
      } else if (timeZone == "M") {
          classTimeBasedOnTimeZone = "1-2 PM";
      } else {
          classTimeBasedOnTimeZone = classTime;
      }
    } else {
      if (timeZone == "P") {
          classTimeBasedOnTimeZone = "1-2 PM";
      } else if (timeZone == "E") {
          classTimeBasedOnTimeZone = "4-5 PM";
      } else if (timeZone == "M") {
          classTimeBasedOnTimeZone = "2-3 PM";
      } else {
          classTimeBasedOnTimeZone = classTime;
      }
    }

    var schedule = studentFN.toString().replace(" ","") + " " + studentLN.toString().replace(" ","") + " has " + activeSheet.getName() + " class at " + classTimeBasedOnTimeZone + " this coming " + classDay + "." 
    var meeting = "Link to join meeting - " + meetingLink

    var worksheet = "";

    if (active == 1) {
        worksheet = activeSheet.getName() + " printout - " + worksheetLink + "\n\n" + "The worksheet section of the packet will be done during the class, while the homework problems should be done bit by bit every day leading up to the next class. Please print these out if possible, so that students can use them." + "\n\n" + "Otherwise, a Google Doc copy can be made and worked on instead of a paper copy." + "\n\n" + "If you have any questions, please let us know."
    } 
    
    if (active == 2) {
        worksheet = "Review Session: Please bring " + activeSheet.getName() + " printouts from past 2 classes";
    } 
    
    if (active == 3) {
        worksheet = studentFN + " will be taking an assessment in class this week. Please print out the following assessment and give it to your student at the beginning of the class." + "\n\n" + worksheetLink;
    } 
    
    if (active == 4) {
        worksheet = "Please bring your students " + activeSheet.getName() + " assessment from last class for review";
    }

    var social = "Connect with us!" + "\n" + "Instagram: @lightbulb_tut (https://www.instagram.com/lightbulb_tut/)" + "\n" + "Facebook: https://m.facebook.com/lbtut/" + "\n" + "Email: contact@lightbulbtutoring.org"
    var footer = "Thank you," + "\n" + "Lightbulb Tutoring" + "\n" + "https://www.lightbulbtutoring.org/"
    
    var mailBody = greeting + "\n\n" + schedule + "\n\n" + meeting + "\n\n" + worksheet + "\n\n"  + footer + "\n\n" + social

    // if student is active
    if (active > 0) {
        // send if Parent eMail is populated
      if (parentEmail.length > 0) {
          // Send if eMail not sent before
          if (dateEmailSent.length == 0) {

            // Send Primary email to Parent
            GmailApp.sendEmail(parentEmail, mailSubject, mailBody);
            emailSentCount = emailSentCount+1;

            // Send Secondary email
            if (secondEmail.length > 0) {
              GmailApp.sendEmail(secondEmail, mailSubject, mailBody);
            }

            // Update sent field with date
            var date = new Date();
            activeSheet.getRange(i,2).setValue(date);

            // Execution Logs
            Logger.log(emailSentCount + " Sent email for "+ studentFN.toString().replace(" ","") + " " + studentLN.toString().replace(" ","") + " to " + parentEmail);

        } 
    }
    }
     

  }
  }
}
