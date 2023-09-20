const scriptProperties = PropertiesService.getScriptProperties()

const getColumnIdByName = (columnName) => {
  const sheet = SpreadsheetApp.getActiveSheet()
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
  for (let index = 0; index < headers.length; index++) {
    if (headers[index] === columnName) {
      return index + 1
    }
  }
  return null
}

const showAlert = (title, message) => {
  const ui = SpreadsheetApp.getUi()
  ui.alert(title, message, ui.ButtonSet.OK)
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
      return showAlert("Heads-up", "Please set description first")
    } else if (date === "") {
      return showAlert("Heads-up", "Please set date first")
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
    const currency = sheet.getRange(row, getColumnIdByName("Currency"))
    if (value !== "") {
      const suppliersSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Suppliers")
      for (let rowId = 2; rowId < suppliersSheet.getLastRow() + 1; rowId++) {
        const supplier = suppliersSheet.getRange(`A${rowId}`).getValue()
        const defaultExpenseCategory = suppliersSheet
          .getRange(`B${rowId}`)
          .getValue()
        const defaultCurrency = suppliersSheet.getRange(`C${rowId}`).getValue()
        if (supplier === value) {
          category.setValue(defaultExpenseCategory)
          currency.setValue(defaultCurrency)
          break
        }
      }
    } else {
      category.setValue("")
      currency.setValue("")
    }
  } else if (
    ["Expenses", "Revenues"].includes(sheetName) &&
    column === getColumnIdByName("Subtotal")
  ) {
    const gst = sheet.getRange(row, getColumnIdByName("GST"))
    const qst = sheet.getRange(row, getColumnIdByName("QST"))
    if (event.range.getValue() !== "") {
      gst.setFormula(`=${event.range.getA1Notation()}*Taxes!B2`)
      qst.setFormula(`=${event.range.getA1Notation()}*Taxes!B3`)
    } else {
      gst.setValue("")
      qst.setValue("")
    }
  }
}
