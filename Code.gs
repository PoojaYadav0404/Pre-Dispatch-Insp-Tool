function copyMainFmt() {
  const NewInspNo = Number(DataSheet.getRange("A2").getValue()) + 1;
  const ThisYear = Utilities.formatDate(new Date(), "GMT+05:30", "yyyy");
  //Logger.log(Test);
  const YearCycle = [];
  if (new Date().getMonth() > 2) {
    YearCycle.push(Number(Utilities.formatDate(new Date(), "GMY+05:30", "yyyy")) + "-" + (Number(ThisYear) + 1));

  }else if (new Date().getMonth() === 2) {
    YearCycle.push((Number(Utilities.formatDate(new Date(), "GMT+05:30", "yyyy")) - 1) + "-" + ThisYear)

  }else if (new Date().getMonth() < 2) {
    YearCycle.push((Number(Utilities.formatDate(new Date(), "GMT+05:30", "yyyy")) - 1) + "-" + ThisYear)
  }

  MainFmt.activate();
  const ActiveSheet = ss.duplicateActiveSheet().setName("IR-" + NewInspNo + "/" + Utilities.formatDate(new Date(), "GMT+05:30", "dd-MMM-yy") + 'Client Name').setTabColor("#6fa8dc").activate();

  //Logger.log(YearCycle)

  ActiveSheet.getRange("K4").setValue([["IR-" + NewInspNo, YearCycle].join("/")]);
  DataSheet.getRange("A2").setValue(NewInspNo);

  MainFmt.hideSheet();
  ss.toast("New format is ready to prepare the inspection report.", "Message!!")

};


function onEdit() {

  if ((!MasterSheets.includes(ActiveSheetName)) && Row===4 && Col === 5 && InspNo != "" && InspDate != "" && CustName != "") {
    const SplitInspNo = InspNo.split("/")[0];
    const ShortDate = Utilities.formatDate(InspDate, "GMT+05:30", "dd-MMM-yyyy");
    const SetSheetName = SplitInspNo + "/" + ShortDate + " " + CustName
    ActiveSheet.setName(SetSheetName)

  } else if ((!MasterSheets.includes(ActiveSheetName)) && Row==5 && Col == 4 && InspNo != "" && InspDate != "" && CustName != "") {
    const SplitInspNo = InspNo.split("/")[0];
    const ShortDate = Utilities.formatDate(InspDate, "GMT+05:30", "dd-MMM-yyyy");
    const SetSheetName = SplitInspNo + "/" + ShortDate + " " + CustName
    ActiveSheet.setName(SetSheetName)

  } 

};




