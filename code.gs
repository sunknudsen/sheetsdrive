const scriptProperties = PropertiesService.getScriptProperties()

const getColumnIdByName = (sheet, columnName) => {
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
      .getRange(row, getColumnIdByName(sheet, "Description"))
      .getValue()
    const date = sheet
      .getRange(row, getColumnIdByName(sheet, "Date"))
      .getValue()
    if (description === "") {
      const error = "Please set description first"
      showAlert("Heads-up", error)
      throw Error(error)
    } else if (date === "") {
      const error = "Please set date first"
      showAlert("Heads-up", error)
      throw Error(error)
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
    const folders = DriveApp.getFolderById(
      scriptProperties.getProperty("folder")
    ).getFoldersByName(sheetFilename)
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
  const menuEntries = [
    { name: "Generate reports", functionName: "generateReports" },
    { name: "Sheetsdrive", functionName: "showSheetsdrive" },
  ]
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

const generateExpenseReport = (currency) => {
  const folder = DriveApp.getFolderById(scriptProperties.getProperty("folder"))
  const sheetFilename = DriveApp.getFileById(
    SpreadsheetApp.getActiveSpreadsheet().getId()
  ).getName()
  const expensesSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenses")
  const expenseReportSheet = SpreadsheetApp.create(
    `${sheetFilename} expense report (${currency})`
  )
  expenseReportSheet.getRange("A1").setFontWeight("bold").setValue("Category")
  expenseReportSheet
    .getRange("B1")
    .setFontWeight("bold")
    .setValue("Percentage used for business activities")
  expenseReportSheet.getRange("C1").setFontWeight("bold").setValue("Amortized")
  expenseReportSheet.getRange("D1").setFontWeight("bold").setValue("Subtotal")
  expenseReportSheet.getRange("E1").setFontWeight("bold").setValue("GST")
  expenseReportSheet.getRange("F1").setFontWeight("bold").setValue("QST")
  const expenseCategoriesSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expense categories")
  for (
    let expenseCategoriesSheetRowId = 2;
    expenseCategoriesSheetRowId < expenseCategoriesSheet.getLastRow() + 1;
    expenseCategoriesSheetRowId++
  ) {
    const expenseCategoryName = expenseCategoriesSheet
      .getRange(`A${expenseCategoriesSheetRowId}`)
      .getValue()
    const expenseCategoryPercentageUsedForBusinessActivities =
      expenseCategoriesSheet
        .getRange(`B${expenseCategoriesSheetRowId}`)
        .getValue()
    const expenseCategoryAmortized = expenseCategoriesSheet
      .getRange(`C${expenseCategoriesSheetRowId}`)
      .getValue()
    expenseReportSheet
      .getRange(`A${expenseCategoriesSheetRowId}`)
      .setValue(expenseCategoryName)
    let subtotal = 0
    let gst = 0
    let qst = 0
    for (
      let expensesRowId = 2;
      expensesRowId < expensesSheet.getLastRow() + 1;
      expensesRowId++
    ) {
      const currentExpenseCategory = expensesSheet
        .getRange(`B${expensesRowId}`)
        .getValue()
      const currentExpenseCurrency = expensesSheet
        .getRange(`E${expensesRowId}`)
        .getValue()
      const currentExpenseRecurrence = expensesSheet
        .getRange(`I${expensesRowId}`)
        .getValue()
      if (
        currentExpenseCategory === expenseCategoryName &&
        currentExpenseCurrency === currency
      ) {
        const currentExpenseSubtotal = expensesSheet
          .getRange(`F${expensesRowId}`)
          .getValue()
        const currentExpenseGst = expensesSheet
          .getRange(`G${expensesRowId}`)
          .getValue()
        const currentExpenseQst = expensesSheet
          .getRange(`H${expensesRowId}`)
          .getValue()
        if (currentExpenseSubtotal !== "") {
          subtotal +=
            currentExpenseRecurrence !== ""
              ? currentExpenseSubtotal * currentExpenseRecurrence
              : currentExpenseSubtotal
        }
        if (currentExpenseGst !== "") {
          gst +=
            currentExpenseRecurrence !== ""
              ? currentExpenseGst * currentExpenseRecurrence
              : currentExpenseGst
        }
        if (currentExpenseQst !== "") {
          qst +=
            currentExpenseRecurrence !== ""
              ? currentExpenseQst * currentExpenseRecurrence
              : currentExpenseQst
        }
      }
    }
    expenseReportSheet
      .getRange(`B${expenseCategoriesSheetRowId}`)
      .setValue(expenseCategoryPercentageUsedForBusinessActivities)
    expenseReportSheet
      .getRange(`C${expenseCategoriesSheetRowId}`)
      .setValue(expenseCategoryAmortized)
    expenseReportSheet
      .getRange(`D${expenseCategoriesSheetRowId}`)
      .setValue(subtotal)
    expenseReportSheet.getRange(`E${expenseCategoriesSheetRowId}`).setValue(gst)
    expenseReportSheet.getRange(`F${expenseCategoriesSheetRowId}`).setValue(qst)
  }
  expenseReportSheet.getDataRange().setFontFamily("Roboto Mono")
  expenseReportSheet
    .getRange("A2:A")
    .setNumberFormat(expenseCategoriesSheet.getRange("A2").getNumberFormat())
  expenseReportSheet
    .getRange("B2:B")
    .setNumberFormat(expenseCategoriesSheet.getRange("B2").getNumberFormat())
  expenseReportSheet
    .getRange("D2:D")
    .setNumberFormat(expensesSheet.getRange("F2").getNumberFormat())
  expenseReportSheet
    .getRange("E2:E")
    .setNumberFormat(expensesSheet.getRange("G2").getNumberFormat())
  expenseReportSheet
    .getRange("F2:F")
    .setNumberFormat(expensesSheet.getRange("H2").getNumberFormat())
  DriveApp.getFileById(expenseReportSheet.getId()).moveTo(folder)
}

const generateReports = () => {
  const currenciesSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Currencies")
  for (
    let currenciesSheetRowId = 2;
    currenciesSheetRowId < currenciesSheet.getLastRow() + 1;
    currenciesSheetRowId++
  ) {
    const currency = currenciesSheet
      .getRange(`A${currenciesSheetRowId}`)
      .getValue()
    generateExpenseReport(currency)
  }
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
  if (
    sheetName === "Expenses" &&
    column === getColumnIdByName(sheet, "Supplier")
  ) {
    const value = event.range.getValue()
    const category = sheet.getRange(row, getColumnIdByName(sheet, "Category"))
    const currency = sheet.getRange(row, getColumnIdByName(sheet, "Currency"))
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
      category.clearContent()
      currency.clearContent()
    }
  } else if (
    sheetName === "Expenses" &&
    column === getColumnIdByName(sheet, "Subtotal")
  ) {
    const supplier = sheet.getRange(row, getColumnIdByName(sheet, "Supplier"))
    const gst = sheet.getRange(row, getColumnIdByName(sheet, "GST"))
    const qst = sheet.getRange(row, getColumnIdByName(sheet, "QST"))
    if (event.range.getValue() !== "") {
      const suppliersSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Suppliers")
      for (let rowId = 2; rowId < suppliersSheet.getLastRow() + 1; rowId++) {
        const name = suppliersSheet.getRange(`A${rowId}`).getValue()
        const taxable = suppliersSheet.getRange(`D${rowId}`).getValue()
        if (name === supplier.getValue() && taxable === "No") {
          return
        }
      }
      gst.setFormula(`=${event.range.getA1Notation()}*Taxes!B2`)
      qst.setFormula(`=${event.range.getA1Notation()}*Taxes!B3`)
    } else {
      gst.clearContent()
      qst.clearContent()
    }
  } else if (
    sheetName === "Revenues" &&
    column === getColumnIdByName(sheet, "Subtotal")
  ) {
    const gst = sheet.getRange(row, getColumnIdByName(sheet, "GST"))
    const qst = sheet.getRange(row, getColumnIdByName(sheet, "QST"))
    if (event.range.getValue() !== "") {
      gst.setFormula(`=${event.range.getA1Notation()}*Taxes!B2`)
      qst.setFormula(`=${event.range.getA1Notation()}*Taxes!B3`)
    } else {
      gst.clearContent()
      qst.clearContent()
    }
  }
}
