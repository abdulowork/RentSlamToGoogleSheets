const configuration = {
  // useActiveSheetInsteadOfSpreadsheetId: false,
  // spreadsheetId: "your-spreadsheet-id",
  // or
  useActiveSheetInsteadOfSpreadsheetId: true,
  addCheckBoxes: true,
  sortByPrice: false,
}

function parseRentSlam() {
  const spreadsheet = configuration.useActiveSheetInsteadOfSpreadsheetId
    ? SpreadsheetApp.getActiveSpreadsheet()
    : SpreadsheetApp.openById(configuration.spreadsheetId)

  if (spreadsheet == null) {
    throw "Spreadsheet is null. Is this script attached to a spreadsheet?"
  }

  const sheet = __prepareDataSheet(spreadsheet)
  const threads = __getRentSlamThreads()

  const lastRow = sheet.getLastRow()
  let parsedMessages = new Set()

  if (lastRow > 1) {
    const existingData = sheet.getRange(2, 2, lastRow - 1, 1).getValues()
    parsedMessages = new Set(existingData.map(row => row[0]))
  }

  for (const thread of threads) {
    for (const message of thread.getMessages()) {
      __parseMessageAndAppendToSheet(message, parsedMessages, sheet)
    }
  }

  if (configuration.sortByPrice) {
    sheet.sort(3)
  }
}

function __getRentSlamThreads() {
  let threads = []
  while (true) {
    const threadsBatch = GmailApp.search("from:property@rentslam.com", threads.length, 500)
    if (threadsBatch.length == 0) { break }
    threads = [...threadsBatch.reverse(), ...threads]
  }
  return threads
}

function __parseMessageAndAppendToSheet(
  message,
  parsedMessages,
  sheet
) {
  // add a character, so that sheet.appendRow doesn't automatically convert messageId to int
  const messageId = "x" + message.getId()
  if (parsedMessages.has(messageId)) { return }

  const date = message.getDate().toISOString()
  const content = message.getBody()

  const prices = Array.from(content.matchAll(/€(\d+)/g)).map(m => m[1])
  const areas = Array.from(content.matchAll(/([\d\?]+)&nbsp/g)).map(m => m[1])
  const rentLinks = Array.from(content.matchAll(/class="mcnButton " title="apply now" href="(.+?)"/g)).map(m => m[1])
  const mapsLinks = Array.from(content.matchAll(/^.*?href="(.*?www\.google\.nl.*?)"/gm)).map(m => m[1].replaceAll(' ', '+'))
  const furnishingCategories = Array.from(content.matchAll(/.*?href=".*?www\.google\.nl.*?<\/a><br>[\r\n]+.*[\r\n]+.*[\r\n]+(.*)<br>/g)).map(m => m[1].toLowerCase())
  const expectedLength = prices.length
  const parameters = [prices, areas, rentLinks, mapsLinks, furnishingCategories]

  for (const [index, parameter] of parameters.entries()) {
    if (parameter.length != expectedLength) {
      Logger.log("Unexpected length in parameter at index: " + String(index) + ", Subject: " + message.getSubject() + ", Date: " + date)
      parameter.splice(0,parameter.length)
      parameter.push(...new Array(expectedLength).fill(String()))
    }
  }

  for (let i = 0; i < expectedLength; i++) {
    const row = [date, messageId]
    for (const parameter of parameters) {
      row.push(parameter[i])
    }

    sheet.appendRow(row)

    if (configuration.addCheckBoxes) {
      const singleRow = 1
      const checkboxCount = 3
      const lastDataColumn = sheet.getLastColumn() - checkboxCount - 1
      sheet.getRange(sheet.getLastRow(), lastDataColumn + 1, singleRow, checkboxCount).insertCheckboxes()
    }

    parsedMessages.add(messageId)
  }
}

function __prepareDataSheet(spreadsheet) {
  const sheetName = "RentSlamData"
  let sheet = spreadsheet.getSheetByName(sheetName)

  if (sheet) { return sheet }

  sheet = spreadsheet.insertSheet(sheetName)

  const checkboxes = configuration.addCheckBoxes ? ["Have seen", "Like", "Applied"] : []
  sheet.appendRow(["Time", "Id", "Price", "Area", "Link", "Map Link", "Furnishing", ...checkboxes, "Notes"])

  if (configuration.addCheckBoxes) {
    __addRule(sheet, "=AND(EQ($I2, TRUE), EQ($J2, FALSE))", "#90ee90")
    __addRule(sheet, "=AND(EQ($I2, FALSE), EQ($H2, TRUE))", "#ffcccb")
    __addRule(sheet, "=EQ($J2, TRUE)", "#ff80ff")
  }

  sheet.setFrozenRows(1)

  return sheet
}

function __addRule(sheet, formula, color) {
  const range = sheet.getRange("A2:Z")

  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground(color)
    .setRanges([range])
    .build()
  const rules = sheet.getConditionalFormatRules()
  rules.push(rule)
  sheet.setConditionalFormatRules(rules)
}
