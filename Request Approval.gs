function requestApproval() {
  if (!MasterSheets.includes(ActiveSheet.getName()) && CustName != ""  && InspDate != "" && InspTime != "" && InspEng != "" && InspNo != "") {

    var InspData = ListSheet.getRange(2, 1, ListSheet.getLastRow(), ListSheet.getLastColumn()).getValues();
    var RecordFound = InspData.filter(r => r[0] == InspNo);

    if (RecordFound.length === 0) {
      var InspReportLink = "https://docs.google.com/spreadsheets/d/1SQm5Y/edit?gid=" + ActiveSheet.getSheetId().toString();
      var url_base = "https://docs.google.com/spreadsheets/d/SS_ID/export?".replace("SS_ID", ss.getId());
      var url_ext = 'export?exportFormat=pdf&format=pdf' //export as pdf/ csv. xls
        // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
        //+
        //('&id=' + ActiveSheet.getSheetId())
        +
        '&size=A4' // paper size
        +
        '&portrait=true' // orientation, false for landscape
        +
        '&scale=2' //1=Normal 100% / 2= Fit to Width / 3=Fit to height / 4=Fit to Page
        //+
        //'&fitw=true' // fit to width, false for actual size
        +
        '&top_margin=0.50' //All four margins must be set!
        +
        '&bottom_margin=0.50' //All four margins must be set!
        +
        '&left_margin=0.50' //All four margins must be set!
        +
        '&right_margin=0.50' //All four margins must be set!
        +
        '&sheetnames=false&printtitle=false&pagenumbers=false' //hide optional headers and footers
        +
        '&gridlines=false' // hide gridlines
        +
        '&fzr=false' // do not repeat row headers (frozen rows) on each page
        + 
        '&gid=';

      var options = {
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        }
      }

      var Response = UrlFetchApp.fetch(url_base + url_ext + ActiveSheet.getSheetId(), options);
      var InspSheetBlob = Response.getBlob().setName(InspNo + " " + CustName + ".pdf");
      var UserId = Session.getActiveUser().getEmail();
      var MailName = MailTemp.getRange("B2").getValue() + "-" + InspEng;
      var To = MailTemp.getRange("B3").getValue();
      var CC = MailTemp.getRange("B4").getValue();
      var Sub = InspNo + MailTemp.getRange("B5").getValue() ;
      var ReplyTo = MailTemp.getRange("B7").getValue() ;
      var Body = MailTemp.getRange("B6").getValue().replace("{report no.}",InspNo).replace("{customer}",CustName).replace("{link}", InspReportLink ) ;

      GmailApp.sendEmail(To, Sub, "", {
        name:MailName, 
        htmlBody: Body, 
        replyTo: UserId, 
        cc: CC, 
        attachments: InspSheetBlob })

      var Arr = [InspNo, InspDate,InspTime, InspEng, CustName, InspReportLink, UserId, "Not approved"];
      ListSheet.getRange(ListSheet.getLastRow() + 1, 1, 1, Arr.length).setValues([Arr]);
      var Ui = SpreadsheetApp.getUi();
      Ui.alert("Approval Mail Status!!"," Inspection Report send succesfully for approval.", Ui.ButtonSet.OK);

    } else {
      var Ui = SpreadsheetApp.getUi();
      Ui.alert("Alert!!", "This Inspection report no. already exist.", Ui.ButtonSet.OK);
    }
  } else {
    var Ui = SpreadsheetApp.getUi();
    Ui.alert("Alert!!", "Either Wrong Sheet or Important details(i.e. Customer Name,Date of Inspection,Time of Inspection, Inspection Engineer, Invoice No., Inspection Report No.) are missing.", Ui.ButtonSet.OK)

  }
  
}
