const ss = SpreadsheetApp.getActiveSpreadsheet();
const MailTemp = ss.getSheetByName("Mail Temp");
const MainFmt = ss.getSheetByName("Insp. Mst Fmt");
const DataSheet = ss.getSheetByName("Customer Data");
const ListSheet = ss.getSheetByName("Insp. List");
const ExtraSheet = ss.getSheetByName("Insp. (2)");
const ActiveSheet = ss.getActiveSheet();
const ActiveSheetName = ActiveSheet.getName();
const InvoiceNo = ActiveSheet.getRange("K6").getValue();
const InspNo = ActiveSheet.getRange("K4").getValue();
const CustName = ActiveSheet.getRange("E4").getValue();
const InspDate = ActiveSheet.getRange("D5").getValue();
const InspTime = ActiveSheet.getRange("D6").getValue();
const InspEng = ActiveSheet.getRange("K5").getValue();
const Row = ActiveSheet.getActiveCell().getRow();
const Col = ActiveSheet.getActiveCell().getColumn();
const MasterSheets = ["Insp. Mst Fmt", "Customer Data", "Insp. List", "Instructions", "Mail Temp"];

function onOpen() {

  const UI = SpreadsheetApp.getUi();
  UI.createMenu("INSPECTION")
    .addItem("New Format", "copyMainFmt")
    .addSeparator()
    .addItem("Add rows", "Row4in1")
    .addSeparator()
    .addItem("Request Approval", "requestApproval")
    .addSeparator()
    .addItem("Approve", "approved")
    .addToUi();

   hideSheets();


}

function hideSheets(){
  var Sheets = ss.getSheets();
  var HiddenSheets = ["Insp. Mst Fmt", "Customer Data", "Mail Temp"];
  Sheets.forEach(function(s){
    if (HiddenSheets.includes(s.getName())){
      s.hideSheet();
    }
  })

}



