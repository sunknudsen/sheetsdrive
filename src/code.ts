const scriptProperties = PropertiesService.getScriptProperties()

interface ColumnIds {
  [name: string]: number
}

const getColumnIds = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
  let columnIds: ColumnIds = {}
  for (let index = 0; index < headers.length; index++) {
    columnIds[headers[index]] = index + 1
  }
  return columnIds
}

const showAlert = (title: string, message: string) => {
  const ui = SpreadsheetApp.getUi()
  ui.alert(title, message, ui.ButtonSet.OK)
}

const slugify = (string: string) => {
  return string
    .toLowerCase()
    .trim()
    .replace(/\s+/g, "-")
    .replace(/[^\w-]+/g, "")
}

const addToDrive = (data: number[], type: string, name: string) => {
  const extension = name.split(".").pop().toLowerCase()
  const sheet = SpreadsheetApp.getActiveSheet()
  const sheetColumnIds = getColumnIds(sheet)
  const selectionRange = sheet.getSelection().getActiveRange()
  if (selectionRange) {
    const row = selectionRange.getRow()
    const description = sheet
      .getRange(row, sheetColumnIds["Description"])
      .getValue()
    const date = sheet.getRange(row, sheetColumnIds["Date"]).getValue()
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
    selectionRange.setFormula(`=HYPERLINK("${file.getUrl()}", "${filename}")`)
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

const generateExpenseReport = (currency: string) => {
  const folder = DriveApp.getFolderById(scriptProperties.getProperty("folder"))
  const sheetFilename = DriveApp.getFileById(
    SpreadsheetApp.getActiveSpreadsheet().getId()
  ).getName()
  const expensesSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenses")
  const expensesSheetColumnIds = getColumnIds(expensesSheet)
  const expenseReportSheet = SpreadsheetApp.create(
    `${sheetFilename} expense report (${currency})`
  )
  expenseReportSheet
    .getRange("A1:F1")
    .setFontWeight("bold")
    .setValues([
      [
        "Category",
        "Percentage used for business activities",
        "Amortized",
        "Subtotal",
        "GST",
        "QST",
      ],
    ])
  const expenseCategoriesSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expense categories")
  const expenseCategoriesSheetColumnIds = getColumnIds(expenseCategoriesSheet)
  for (
    let expenseCategoriesSheetRowId = 2;
    expenseCategoriesSheetRowId < expenseCategoriesSheet.getLastRow() + 1;
    expenseCategoriesSheetRowId++
  ) {
    const expenseCategoryName = expenseCategoriesSheet
      .getRange(
        expenseCategoriesSheetRowId,
        expenseCategoriesSheetColumnIds["Name"]
      )
      .getValue()
    const expenseCategoryPercentageUsedForBusinessActivities =
      expenseCategoriesSheet
        .getRange(
          expenseCategoriesSheetRowId,
          expenseCategoriesSheetColumnIds[
            "Percentage used for business activities"
          ]
        )
        .getValue()
    const expenseCategoryAmortized = expenseCategoriesSheet
      .getRange(
        expenseCategoriesSheetRowId,
        expenseCategoriesSheetColumnIds["Amortized"]
      )
      .getValue()
    let subtotal = 0
    let gst = 0
    let qst = 0
    for (
      let expensesRowId = 2;
      expensesRowId < expensesSheet.getLastRow() + 1;
      expensesRowId++
    ) {
      const currentExpenseCategory = expensesSheet
        .getRange(expensesRowId, expensesSheetColumnIds["Category"])
        .getValue()
      const currentExpenseCurrency = expensesSheet
        .getRange(expensesRowId, expensesSheetColumnIds["Currency"])
        .getValue()
      const currentExpenseRecurrence = expensesSheet
        .getRange(expensesRowId, expensesSheetColumnIds["Recurrence"])
        .getValue()
      if (
        currentExpenseCategory === expenseCategoryName &&
        currentExpenseCurrency === currency
      ) {
        const currentExpenseSubtotal = expensesSheet
          .getRange(expensesRowId, expensesSheetColumnIds["Subtotal"])
          .getValue()
        const currentExpenseGst = expensesSheet
          .getRange(expensesRowId, expensesSheetColumnIds["GST"])
          .getValue()
        const currentExpenseQst = expensesSheet
          .getRange(expensesRowId, expensesSheetColumnIds["QST"])
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
      .getRange(
        `A${expenseCategoriesSheetRowId}:F${expenseCategoriesSheetRowId}`
      )
      .setValues([
        [
          expenseCategoryName,
          expenseCategoryPercentageUsedForBusinessActivities,
          expenseCategoryAmortized,
          subtotal,
          gst,
          qst,
        ],
      ])
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

const onEdit = (event: GoogleAppsScript.Events.SheetsOnEdit) => {
  const row = event.range.getRow()
  const column = event.range.getColumn()
  const sheet = SpreadsheetApp.getActiveSheet()
  const sheetName = sheet.getName()
  const sheetColumnIds = getColumnIds(sheet)
  if (sheetName === "Expenses" && column === sheetColumnIds["Supplier"]) {
    const value = event.range.getValue()
    const category = sheet.getRange(row, sheetColumnIds["Category"])
    const currency = sheet.getRange(row, sheetColumnIds["Currency"])
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
    column === sheetColumnIds["Subtotal"]
  ) {
    const supplier = sheet.getRange(row, sheetColumnIds["Supplier"])
    const gst = sheet.getRange(row, sheetColumnIds["GST"])
    const qst = sheet.getRange(row, sheetColumnIds["QST"])
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
    column === sheetColumnIds["Subtotal"]
  ) {
    const gst = sheet.getRange(row, sheetColumnIds["GST"])
    const qst = sheet.getRange(row, sheetColumnIds["QST"])
    if (event.range.getValue() !== "") {
      gst.setFormula(`=${event.range.getA1Notation()}*Taxes!B2`)
      qst.setFormula(`=${event.range.getA1Notation()}*Taxes!B3`)
    } else {
      gst.clearContent()
      qst.clearContent()
    }
  }
}
