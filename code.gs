const scriptProperties = PropertiesService.getScriptProperties()

const slugify = (string) => {
  return string
    .toLowerCase()
    .trim()
    .replace(/\s+/g, "-")
    .replace(/[^\w-]+/g, "")
}

const addToDrive = (data, type, name) => {
  const extension = name.split(".").pop().toLowerCase()
  const sheet = SpreadsheetApp.getActiveSheet()
  const selection = sheet.getActiveSelection()
  if (selection) {
    const row = selection.getRow()
    const column = selection.getColumn()
    const title = sheet.getRange(row, column - 2).getValue()
    const date = sheet.getRange(row, column - 1).getValue()
    if (title === "") {
      throw new Error("Please set title first")
    } else if (date === "") {
      throw new Error("Please set date first")
    }
    const formattedDate = Utilities.formatDate(
      date,
      "America/Montreal",
      "yyyy-MM-dd"
    )
    const structuredName = `${slugify(title)}-${formattedDate}.${extension}`
    const blob = Utilities.newBlob(data, type, structuredName)
    const file = DriveApp.getFolderById(
      scriptProperties.getProperty("folderId")
    ).createFile(blob)
    selection.setFormula(`=HYPERLINK("${file.getUrl()}", "${structuredName}")`)
    return
  }
}

function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
  const menuEntries = [{ name: "Sheetsdrive", functionName: "showSheetsdrive" }]
  sheet.addMenu("Custom utilities", menuEntries)
}

function showSheetsdrive() {
  const template = HtmlService.createTemplateFromFile("upload")
  template.webAppUrl = scriptProperties.getProperty("webAppUrl")
  const html = template
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle("Sheetsdrive")
  // SpreadsheetApp.getUi().showModalDialog(html, "Sheetsdrive")
  SpreadsheetApp.getUi().showSidebar(html)
}
