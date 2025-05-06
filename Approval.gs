function approved() {
  var UserId = Session.getActiveUser().getEmail();
  var AppBy = ActiveSheet.getRange(ActiveSheet.getLastRow(), 11).getValue();

  if (!MasterSheets.includes(ActiveSheetName) && AppBy != "" && (UserId == "") && CustName != ""  && InspDate != "" && InspTime != "" && InspEng != "" && InspNo != "") {

    var InspArr = (!ListSheet.getLastRow() < 1) ? ListSheet.getRange(2, 1, ListSheet.getLastRow() - 1, 1).getValues().map(insp => { return insp.toString() }) : [];
    var Row = InspArr.indexOf(InspNo.toString()) + 2;
    Logger.log(InspArr);
    Logger.log(Row);

    if(Row > 1){
      var Ui = SpreadsheetApp.getUi();
      var Response = Ui.prompt("Carton Details!!", "How many no. of cartons used for this order(in no's like 10, etc.)?", Ui.ButtonSet.OK_CANCEL);

      if (Response.getSelectedButton() === Ui.Button.OK) {
        var Ans = Response.getResponseText();
        Logger.log(Ans);

        var FolderId = "5Sz93MYApWJ"
        var MasterFormatId = "1i546XQOZ";
        var MainDrive = DriveApp.getFolderById(FolderId);
        try {
          var CustomerFdr = MainDrive.getFoldersByName(CustName).next();
        } catch (e) {
          var CustomerFdr = MainDrive.createFolder(CustName);
        }

        var NewSheetId = DriveApp.getFileById(MasterFormatId).makeCopy(CustomerFdr).setName(InspNo + "//" + CustName).getId();
        var NewSheet = SpreadsheetApp.openById(NewSheetId);
        ActiveSheet.copyTo(NewSheet);

        var sheet = NewSheet.getSheetByName('Sheet1');
        NewSheet.deleteSheet(sheet);

        var FinalNewSheet = NewSheet.getSheets()[0].setName(InspNo + "//" + CustName).activate();
        FinalNewSheet.getRange(1, 1, FinalNewSheet.getLastRow(), FinalNewSheet.getLastColumn()).activate();
        FinalNewSheet.getActiveRangeList().setFontFamily('Calibri');

        //Delete data validations
        FinalNewSheet.getRange("E4").clearDataValidations();
        FinalNewSheet.getRange("K5").clearDataValidations();
        FinalNewSheet.getRange(9, 2, FinalNewSheet.getLastRow() - 1 - 8, 2).clearDataValidations();
        FinalNewSheet.getRange(9, 7, FinalNewSheet.getLastRow() - 1 - 8, 9).clearDataValidations();
        FinalNewSheet.getRange(FinalNewSheet.getLastRow(), 11).clearDataValidations();

        // ceating PDF for mail
        var url_base = "https://docs.google.com/spreadsheets/d/SS_ID/export?".replace("SS_ID", NewSheetId);
        var url_ext = 'export?exportFormat=pdf&format=pdf' //export as pdf/ csv. xls
          // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
          // following parameters are optional...
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

        var Response = UrlFetchApp.fetch(url_base + url_ext + FinalNewSheet.getSheetId(), options);
        var InspSheetBlob = Response.getBlob().setName(InspNo + " " + CustName + ".pdf");
        var To = MailTemp.getRange("E3").getValue();
        var CC = MailTemp.getRange("E4").getValue();
        var Sub = InspNo + MailTemp.getRange("E5").getValue();
        var RepBody = MailTemp.getRange("E6").getValue().replace("{carton}",Ans);
        var ReplyTo = MailTemp.getRange("E7").getValue();
        //GmailApp.sendEmail(To, Sub, "", { cc: CC, htmlBody: RepBody, attachments: InspSheetBlob, replyTo: ReplyTo });

        //GmailApp.getThreadById(ThreadId.toString()).replyAll("", { htmlBody: RepBody, attachments: InspSheetBlob });


        var InspArr = (!ListSheet.getLastRow() < 1) ? ListSheet.getRange(2, 1, ListSheet.getLastRow() - 1, 1).getValues().map(insp => { return insp.toString() }) : [];
        var Row = InspArr.indexOf(InspNo.toString()) + 2;
        Logger.log(Row)
        ListSheet.getRange(Row, 6).setValue([NewSheet.getUrl()]);
        ListSheet.getRange(Row, 8).setValue(["Approved"]);
        ListSheet.getRange(Row, 9).setValue([new Date()]);
        ListSheet.getRange(Row, 10).setValue([Ans]);
        ss.deleteActiveSheet();

        ss.toast(" Inspection report created succesfully and Details are also shared with the Sales Team #sales@bajato.com.", "Successful!!");

      } else {
        Ui.alert("Error Message!!", "Please try again and enter No. of Cartons then click Ok", Ui.ButtonSet.OK)
      };
    } else{ var Ui = SpreadsheetApp.getUi();
      Ui.alert("###ERROR###", "Till now, this inspection report is not requested for approval.", Ui.ButtonSet.OK )

    }

  } else { var Ui = SpreadsheetApp.getUi();
    Ui.alert("###ERROR###", "Either wrong sheet or important details are missing like 'Customer Name', 'Insp. Report No.', 'Date of Insp.', 'Insp. Engineer', 'Time of Insp.', 'Invoice No.'", Ui.ButtonSet.OK ) }

};
