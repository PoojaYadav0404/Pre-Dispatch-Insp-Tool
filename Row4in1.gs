function Row4in1() {
  if (MasterSheets.includes(ActiveSheetName)) {
    const Ui = SpreadsheetApp.getUi();
    Ui.alert("Warning!!", "Adding row in wrong sheet.", Ui.ButtonSet.OK);


  } else if ((!MasterSheets.includes(ActiveSheetName)) && ActiveSheet.getRange("E4").getValue() == "") {
    const Ui = SpreadsheetApp.getUi();
    Ui.alert("Warning!!", "Please fill customer name.", Ui.ButtonSet.OK);

  } else if ((!MasterSheets.includes(ActiveSheetName)) && ActiveSheet.getRange("E4").getValue() != "") {
    const AddRowPosition = ActiveSheet.getLastRow() - 1;
    ActiveSheet.insertRowsAfter(AddRowPosition, 4);

    //.............copy format only
    ActiveSheet.getRange(ActiveSheet.getLastRow()-8, 1, 4, 15).copyTo(ActiveSheet.getRange(ActiveSheet.getLastRow()-4, 1, 4, 15), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false)
    ActiveSheet.getRange(ActiveSheet.getLastRow()-4, 1, 4, 4).clearContent();
    ActiveSheet.getRange(ActiveSheet.getLastRow()-4, 7, 4, 9).clearContent();

    //line formatting on approved by
    //ActiveSheet.getRange(ActiveSheet.getLastRow(), 1, 1, 15).setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    ss.toast("New rows added successfully.", "Message!!");

  }

}


