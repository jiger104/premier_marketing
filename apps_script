var ui = SpreadsheetApp.getUi();


function onOpen(e) {

  // Create menu options

  ui.createAddonMenu()
    .addItem("PM Data Import", "buildModalPM")
    .addToUi();
};


function buildModalPM() {
  //Call the HTML file and set the width and height
  var html = HtmlService.createHtmlOutputFromFile("Importer Menu")
    .setWidth(400)
    .setHeight(200);

  //Display the dialog
  var dialog = ui.showModalDialog(html, "Provide Info Below");

};

function runQBOImport () {
  // Copy over daily qbo ad spend report & run qbo import functions
  newQBOData = SpreadsheetApp.openById("1hqL4UvGSyShJ3bhzIm-9uLzXIaLcWx-TD52WeGcHHe4").getActiveSheet().getDataRange().getValues()
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QBO").clear()
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QBO").getRange(1,1,newQBOData.length,7).setValues(newQBOData)
  currentMonth = new Date()
  currentMonth.setMonth(currentMonth.getMonth()+1)
  currentMonth = Utilities.formatDate(currentMonth,"GMT+1:00","YYYY-MM")
  runImporter({ importType: "QBO", importDate: currentMonth })
}

function parseGoogleData(timePeriod, sheetNames) {
  googleData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Google")


  lastRow = googleData.getLastRow()
  googleClients = googleData.getRange(4, 1, lastRow, 4).getValues().filter(String)
  unmatchedClients = []
  // Iterate through google ad client list and find matching tab in workbook
  for (i in googleClients) {
    clientName = googleClients[i][0]
    billedCost = googleClients[i][3]
    if (sheetNames.includes(clientName)) {
      console.log("Found Google Client Match: ", clientName)
      // If matching tab found, check if there is pre-existing month or new one needs to be created based on the time period we are importing
      matchedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(clientName)
      matchedSheet.getRange("A:A").setNumberFormat("mm/dd/yyyy")
      matchedSheetDates = matchedSheet.getRange(2, 1, matchedSheet.getLastRow(), 1).getValues().map(date => Utilities.formatDate(new Date(date), "GMT+1:00", "MMM YY"))
      matchedDateRow = matchedSheetDates.includes(timePeriod) ? matchedSheetDates.indexOf(timePeriod) + 2 : null
      if (matchedDateRow !== null) {
        matchedSheet.getRange(matchedDateRow, 4).setValue(billedCost)
      }
      else {
        matchedSheet.getRange(matchedSheet.getLastRow() + 1, 1).setValue(timePeriod)
        matchedSheet.getRange(matchedSheet.getLastRow(), 4).setValue(billedCost)
      }


    }
    else {
      unmatchedClients.push(clientName)
    }
  }
  return unmatchedClients.filter(String)
}

function parseFacebookData(timePeriod, sheetNames) {
  facebookData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Facebook")

  lastRow = facebookData.getLastRow()
  facebookClients = facebookData.getRange(3, 1, lastRow, 6).getValues().filter(String)
  unmatchedClients = []
  headers = facebookData.getRange()
  // Iterate through facebook ad client list and find matching tab in workbook
  for (i in facebookClients) {
    clientName = facebookClients[i][0]
    billedCost = facebookClients[i][5]
    if (sheetNames.includes(clientName)) {
      console.log("Found FB Client Match: ", clientName)
      // If matching tab found, check if there is pre-existing month or new one needs to be created based on the time period we are importing
      matchedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(clientName)
      matchedSheet.getRange("H:H").setNumberFormat("mm/dd/yyyy")
      matchedSheetDates = matchedSheet.getRange(`H2:H${matchedSheet.getLastRow()}`).getValues().map(date => Utilities.formatDate(new Date(date), "GMT+1:00", "MMM YY"))
      matchedDateRow = matchedSheetDates.includes(timePeriod) ? matchedSheetDates.indexOf(timePeriod) + 2 : null
      if (matchedDateRow !== null) {
        matchedSheet.getRange(matchedDateRow, 11).setValue(billedCost)
      }
      else {
        matchedSheet.getRange(matchedSheetDates.length + 2, 8).setValue(timePeriod)
        matchedSheet.getRange(matchedSheetDates.length + 2, 11).setValue(billedCost)
        matchedSheet.getRange(matchedSheetDates.length + 2,10).setFormula()
      }


    }
    else {
      unmatchedClients.push(clientName)
    }
  }
  return unmatchedClients.filter(String)
}

function preProcessQBOData() {
  // Clean up raw qbo class spend report & objects to summarize spend by class/client
  qboDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QBO")
  qboBudgetSummary = {}
  if (qboDataSheet.getRange(1, 1).getValues()[0][0] !== "Date") {
    qboDataSheet.deleteRows(1, 4)
    qboDataSheet.deleteColumns(1)
    qboDataSheet.deleteRows(2, 2)
    lastRow = qboDataSheet.getRange("A:A").getValues().filter(String).length
    qboDataSheet.deleteRows(lastRow + 1, 10)
  }
  lastRow = qboDataSheet.getRange("A:A").getValues().filter(String).length
  qboSpendData = qboDataSheet.getRange(2, 1, lastRow, 6).getValues()
  // Iterate through rows and create object map of client + spend amount
  for (i in qboSpendData) {
    clientName = qboSpendData[i][5] !== null ? qboSpendData[i][5].split(":").splice(-1)[0] : null
    budgetAmount = qboSpendData[i][4]
    budgetType = qboSpendData[i][3]
    if (budgetType.includes("Google")) {
      budgetType = "Google"
    }
    else if (budgetType.includes("Social") || budgetType.includes("Facebook") || budgetType.includes("FB")) {
      budgetType = "Facebook"
    }
    else {
      budgetType = null
    }
    if (clientName in qboBudgetSummary) {
      qboBudgetSummary[clientName]['Facebook'] = budgetType == "Facebook" ? qboBudgetSummary[clientName]['Facebook'] + budgetAmount : qboBudgetSummary[clientName]['Facebook']
      qboBudgetSummary[clientName]['Google'] = budgetType == "Google" ? qboBudgetSummary[clientName]['Google'] + budgetAmount : qboBudgetSummary[clientName]['Google']
    }
    else if (!(clientName in qboBudgetSummary) && clientName !== '') {
      facebookBudget = budgetType == "Facebook" ? budgetAmount : 0
      googleBudget = budgetType == "Google" ? budgetAmount : 0
      qboBudgetSummary[clientName] = { "Facebook": facebookBudget, "Google": googleBudget }
    }
  }
  return qboBudgetSummary
}

function parseQBOData(timePeriod, sheetNames, qboBudgetSummary) {
  unmatchedClients = []
  // Iterate through qbo client list and find matching tab in workbook
  for (clientName in qboBudgetSummary) {
    if (sheetNames.includes(clientName)) {
      console.log("Found QBO Client Match: ", clientName)
      // If matching tab found, check if there is pre-existing month or new one needs to be created based on the time period we are importing
      matchedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(clientName)
      matchedSheet.getRange("A:A").setNumberFormat("mm/dd/yyyy")
      matchedSheet.getRange("H:H").setNumberFormat("mm/dd/yyyy")
      matchedSheetDatesGoogle = matchedSheet.getRange("A:A").getValues().map(date => Utilities.formatDate(new Date(date), "GMT+1:00", "MMM YY"))
      matchedSheetDatesFacebook = matchedSheet.getRange("H:H").getValues().map(date => Utilities.formatDate(new Date(date), "GMT+1:00", "MMM YY"))
      matchedDateRowGoogle = matchedSheetDatesGoogle.includes(timePeriod) ? matchedSheetDatesGoogle.indexOf(timePeriod) : null
      matchedDateRowFacebook = matchedSheetDatesFacebook.includes(timePeriod) ? matchedSheetDatesFacebook.indexOf(timePeriod) : null
      if (matchedDateRowGoogle !== null) {
        matchedSheet.getRange(matchedDateRowGoogle + 1, 2).setValue(qboBudgetSummary[clientName]["Google"])
      }
      if (matchedDateRowFacebook !== null) {
        matchedSheet.getRange(matchedDateRowFacebook + 1, 9).setValue(qboBudgetSummary[clientName]["Facebook"])
      }
      if (matchedDateRowGoogle == null) {
        // If matching month not found, create new month rows under Google and FB columns
        matchedSheet.getRange(matchedSheet.getLastRow() + 1, 1).setValue(timePeriod)
        matchedSheet.getRange(matchedSheet.getLastRow(), 2).setValue(qboBudgetSummary[clientName]["Google"])
      }
      if (matchedDateRowFacebook == null && qboBudgetSummary[clientName]["Facebook"] > 0) {
        // If matching month not found, create new month rows under Google and FB columns
        matchedSheet.getRange(matchedSheet.getLastRow(), 8).setValue(timePeriod)
        matchedSheet.getRange(matchedSheet.getLastRow(), 9).setValue(qboBudgetSummary[clientName]["Facebook"])
      }


    }
    else {
      unmatchedClients.push(clientName)
    }
  }
  return unmatchedClients.filter(String)
}

function getSheetNames() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheets()
    .map(s => s.getName())
    ;
}

function testFunc() {
  values = { importType: "Facebook", importDate: "2021-07" }
  runImporter(values)
}

function runImporter(values) {
  sheetNames = getSheetNames()
  timePeriod = Utilities.formatDate(new Date(values.importDate), "GMT+1:00", "MMM YY")
  try {
    switch (values.importType) {

      case "Google":
        unmatchedClients = parseGoogleData(timePeriod, sheetNames)
        break

      case "Facebook":
        unmatchedClients = parseFacebookData(timePeriod, sheetNames)
        break


      case "QBO":
        unmatchedClients = parseQBOData(timePeriod, sheetNames, preProcessQBOData())
        break
    }
    console.log(`${values.importType} Import Sucessfull`)
    logAttempt(values.importType, timePeriod, unmatchedClients)
    return
  }
  catch (error) {
    console.log(`Error Occured Running ${values.importType} Import`)
    console.log(error)
    logAttempt(values.importType, timePeriod, error = error)
    throw Exception
  }
}

function logAttempt(importType, timePeriod, unmatchedClients = null, error = null) {
  ss = SpreadsheetApp.getActiveSpreadsheet()
  logs = ss.getSheetByName('Log')
  lastRow = logs.getLastRow() + 1
  if (error == null) {
    logs.getRange(lastRow, 1).setValue(Utilities.formatDate(new Date(), "GMT-6:00", "YYYY-MM-dd HH:mm:ss"))
    logs.getRange(lastRow, 2).setValue(importType)
    logs.getRange(lastRow, 3).setValue(timePeriod)
    logs.getRange(lastRow, 4).setValue("Sucess")
    logs.getRange(lastRow, 6).setValue("N/A")
    if (unmatchedClients.length > 0) {
      logs.getRange(lastRow, 5).setValue(unmatchedClients.toString())
    }
    logs.activate()
  }
  else {
    logs.getRange(lastRow, 1).setValue(Utilities.formatDate(new Date(), "GMT-6:00", "YYYY-MM-dd HH:mm:ss"))
    logs.getRange(lastRow, 2).setValue(importType)
    logs.getRange(lastRow, 3).setValue(timePeriod)
    logs.getRange(lastRow, 4).setValue("Failure")
    logs.getRange(lastRow, 6).setValue(error)
    logs.activate()
  }
}

