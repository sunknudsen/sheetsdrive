const scriptProperties = PropertiesService.getScriptProperties()

const getColumnIdByName = (columnName) => {
  const sheet = SpreadsheetApp.getActiveSheet()
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === columnName) {
      return i + 1
    }
  }
  return null
}

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
    const description = sheet
      .getRange(row, getColumnIdByName("Description"))
      .getValue()
    const date = sheet.getRange(row, getColumnIdByName("Date")).getValue()
    if (description === "") {
      throw new Error("Please set description first")
    } else if (date === "") {
      throw new Error("Please set date first")
    }
    const formattedDate = Utilities.formatDate(
      date,
      "America/Montreal",
      "yyyy-MM-dd"
    )
    const sheetFilename = DriveApp.getFileById(
      SpreadsheetApp.getActiveSpreadsheet().getId()
    ).getName()
    const filename = `${formattedDate}-${slugify(description)}.${extension}`
    const blob = Utilities.newBlob(data, type, filename)
    const folders = DriveApp.getRootFolder().getFoldersByName(sheetFilename)
    let folderId = null
    if (folders.hasNext()) {
      folderId = folders.next().getId()
    } else {
      folderId = DriveApp.getRootFolder().createFolder(sheetFilename).getId()
    }
    const file = DriveApp.getFolderById(folderId).createFile(blob)
    selection.setFormula(`=HYPERLINK("${file.getUrl()}", "${filename}")`)
    return
  }
}

const onOpen = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
  const menuEntries = [{ name: "Sheetsdrive", functionName: "showSheetsdrive" }]
  sheet.addMenu("Custom utilities", menuEntries)
}

const showSheetsdrive = () => {
  const template = HtmlService.createTemplateFromFile("upload")
  template.webAppUrl = scriptProperties.getProperty("webAppUrl")
  const html = template
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle("Sheetsdrive")
  SpreadsheetApp.getUi().showSidebar(html)
}

const onEdit = (event) => {
  const sheet = SpreadsheetApp.getActiveSheet()
  const sheetName = sheet.getName()
  if (
    event.range.rowStart !== event.range.rowEnd &&
    event.range.columnStart !== event.range.columnEnd
  ) {
    // More than one cell selected, stop
    return
  }
  const row = event.range.rowStart
  const column = event.range.columnStart
  if (sheetName === "Expenses" && column === getColumnIdByName("Supplier")) {
    const value = event.range.getValue()
    const category = sheet.getRange(row, getColumnIdByName("Category"))
    if (value !== "") {
      const variablesSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Variables")
      for (let rowId = 3; rowId < variablesSheet.getLastRow() + 1; rowId++) {
        const supplier = variablesSheet.getRange(`F${rowId}`).getValue()
        const defaultExpenseCategory = variablesSheet
          .getRange(`G${rowId}`)
          .getValue()
        if (supplier === "") {
          break
        } else if (supplier === value) {
          category.setValue(defaultExpenseCategory)
          break
        }
      }
    } else {
      category.setValue("")
    }
  } else if (
    ["Expenses", "Revenues"].includes(sheetName) &&
    column === getColumnIdByName("Subtotal")
  ) {
    const gst = sheet.getRange(row, getColumnIdByName("GST"))
    const qst = sheet.getRange(row, getColumnIdByName("QST"))
    if (event.range.getValue() !== "") {
      gst.setFormula(`=${event.range.getA1Notation()}*Variables!J3`)
      qst.setFormula(`=${event.range.getA1Notation()}*Variables!J4`)
    } else {
      gst.setValue("")
      qst.setValue("")
    }
  }
}
