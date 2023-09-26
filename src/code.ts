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
    const values = sheet
      .getRange(row, 1, 1, selectionRange.getLastColumn())
      .getValues()
    const supplier = values[0][sheetColumnIds["Supplier"] - 1]
    const description = values[0][sheetColumnIds["Description"] - 1]
    const date = values[0][sheetColumnIds["Date"] - 1]
    if (description === "") {
      const range = sheet.getRange(row, sheetColumnIds["Description"])
      sheet.setActiveSelection(range)
      SpreadsheetApp.flush()
      const error = `Please set description first at ${range.getA1Notation()}`
      showAlert("Heads-up", error)
      throw Error(error)
    } else if (date === "") {
      const range = sheet.getRange(row, sheetColumnIds["Date"])
      sheet.setActiveSelection(range)
      SpreadsheetApp.flush()
      const error = `Please set date first at ${range.getA1Notation()}`
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
    const filename = `${formattedDate}-${slugify(supplier)}-${slugify(
      description
    )}.${extension}`
    const blob = Utilities.newBlob(data, type, filename)
    const folders = DriveApp.getFolderById(
      scriptProperties.getProperty("folder")
    ).getFoldersByName(sheetFilename)
    let folderId = null
    if (folders.hasNext()) {
      folderId = folders.next().getId()
    } else {
      folderId = DriveApp.getFolderById(scriptProperties.getProperty("folder"))
        .createFolder(sheetFilename)
        .getId()
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

const generateExpenseReport = (currency: string, decimalPlace: number) => {
  const folder = DriveApp.getFolderById(scriptProperties.getProperty("folder"))
  const sheetFilename = DriveApp.getFileById(
    SpreadsheetApp.getActiveSpreadsheet().getId()
  ).getName()
  const expensesSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenses")
  const expensesSheetValues = expensesSheet
    .getRange(1, 1, expensesSheet.getLastRow(), expensesSheet.getLastColumn())
    .getValues()
  const expensesSheetHeaders = expensesSheetValues[0]
  const expenseReportSheet = SpreadsheetApp.create(
    `${sheetFilename} expense report (${currency})`
  )
  expenseReportSheet
    .getRange("A1:G1")
    .setFontWeight("bold")
    .setValues([
      [
        "Category",
        "Percentage used for business activities",
        "Capital expense",
        "Subtotal",
        "Subtotal (CAD)",
        "GST",
        "QST",
      ],
    ])
  const expenseCategoriesSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expense categories")
  const expenseCategoriesSheetValues = expenseCategoriesSheet
    .getRange(
      1,
      1,
      expenseCategoriesSheet.getLastRow(),
      expenseCategoriesSheet.getLastColumn()
    )
    .getValues()
  const expenseCategoriesSheetHeaders = expenseCategoriesSheetValues[0]
  for (
    let expenseCategoriesSheetValuesIndex = 1;
    expenseCategoriesSheetValuesIndex < expenseCategoriesSheetValues.length;
    expenseCategoriesSheetValuesIndex++
  ) {
    const expenseCategoryName =
      expenseCategoriesSheetValues[expenseCategoriesSheetValuesIndex][
        expenseCategoriesSheetHeaders.indexOf("Name")
      ]
    const expenseCategoryPercentageUsedForBusinessActivities =
      expenseCategoriesSheetValues[expenseCategoriesSheetValuesIndex][
        expenseCategoriesSheetHeaders.indexOf(
          "Percentage used for business activities"
        )
      ]
    const expenseCategoryCapitalExpense =
      expenseCategoriesSheetValues[expenseCategoriesSheetValuesIndex][
        expenseCategoriesSheetHeaders.indexOf("Capital expense")
      ]
    let subtotal = 0
    let subtotalCad = 0
    let gst = 0
    let qst = 0
    for (
      let expensesSheetValuesIndex = 1;
      expensesSheetValuesIndex < expensesSheetValues.length;
      expensesSheetValuesIndex++
    ) {
      const currentExpenseCategory =
        expensesSheetValues[expensesSheetValuesIndex][
          expensesSheetHeaders.indexOf("Category")
        ]
      const currentExpenseCurrency =
        expensesSheetValues[expensesSheetValuesIndex][
          expensesSheetHeaders.indexOf("Currency")
        ]
      if (
        currentExpenseCategory === expenseCategoryName &&
        currentExpenseCurrency === currency
      ) {
        const currentExpenseSubtotal =
          expensesSheetValues[expensesSheetValuesIndex][
            expensesSheetHeaders.indexOf("Subtotal")
          ]
        const currentExpenseSubtotalCad =
          expensesSheetValues[expensesSheetValuesIndex][
            expensesSheetHeaders.indexOf("Subtotal (CAD)")
          ]
        const currentExpenseGst =
          expensesSheetValues[expensesSheetValuesIndex][
            expensesSheetHeaders.indexOf("GST")
          ]
        const currentExpenseQst =
          expensesSheetValues[expensesSheetValuesIndex][
            expensesSheetHeaders.indexOf("QST")
          ]
        const currentExpenseRecurrence =
          expensesSheetValues[expensesSheetValuesIndex][
            expensesSheetHeaders.indexOf("Recurrence")
          ]
        if (currentExpenseSubtotal !== "") {
          subtotal +=
            currentExpenseRecurrence !== ""
              ? currentExpenseSubtotal * currentExpenseRecurrence
              : currentExpenseSubtotal
        }
        if (currentExpenseSubtotal !== "" && currentExpenseSubtotalCad === "") {
          const range = expensesSheet.getRange(
            expensesSheetValuesIndex + 1,
            expensesSheetHeaders.indexOf("Subtotal (CAD)") + 1
          )
          expensesSheet.setActiveSelection(range)
          SpreadsheetApp.flush()
          throw new Error(`Missing data at ${range.getA1Notation()}`)
        }
        if (currentExpenseSubtotalCad !== "") {
          subtotalCad +=
            currentExpenseRecurrence !== ""
              ? currentExpenseSubtotalCad * currentExpenseRecurrence
              : currentExpenseSubtotalCad
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
        `A${expenseCategoriesSheetValuesIndex + 1}:G${
          expenseCategoriesSheetValuesIndex + 1
        }`
      )
      .setValues([
        [
          expenseCategoryName,
          expenseCategoryPercentageUsedForBusinessActivities,
          expenseCategoryCapitalExpense,
          subtotal,
          subtotalCad,
          gst,
          qst,
        ],
      ])
  }
  expenseReportSheet.getDataRange().setFontFamily("Roboto Mono")
  expenseReportSheet.getRange("A2:A").setNumberFormat("@")
  expenseReportSheet.getRange("B2:B").setNumberFormat("0.00%")
  expenseReportSheet.getRange("C2:C").setNumberFormat("@")
  expenseReportSheet
    .getRange("D2:D")
    .setNumberFormat(`0.${"0".repeat(decimalPlace)}`)
  expenseReportSheet
    .getRange("E2:E")
    .setNumberFormat(`0.${"0".repeat(decimalPlace)}`)
  expenseReportSheet
    .getRange("F2:F")
    .setNumberFormat(`0.${"0".repeat(decimalPlace)}`)
  expenseReportSheet
    .getRange("G2:G")
    .setNumberFormat(`0.${"0".repeat(decimalPlace)}`)
  DriveApp.getFileById(expenseReportSheet.getId()).moveTo(folder)
}

const generateRevenueReport = (currency: string, decimalPlace: number) => {
  const folder = DriveApp.getFolderById(scriptProperties.getProperty("folder"))
  const sheetFilename = DriveApp.getFileById(
    SpreadsheetApp.getActiveSpreadsheet().getId()
  ).getName()
  const revenuesSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Revenues")

  const revenuesSheetValues = revenuesSheet
    .getRange(1, 1, revenuesSheet.getLastRow(), revenuesSheet.getLastColumn())
    .getValues()
  const revenuesSheetHeaders = revenuesSheetValues[0]
  const revenueReportSheet = SpreadsheetApp.create(
    `${sheetFilename} revenue report (${currency})`
  )
  revenueReportSheet
    .getRange("A1:D1")
    .setFontWeight("bold")
    .setValues([["Subtotal", "Subtotal (CAD)", "GST", "QST"]])
  let subtotal = 0
  let subtotalCad = 0
  let gst = 0
  let qst = 0
  for (
    let revenuesSheetValuesIndex = 1;
    revenuesSheetValuesIndex < revenuesSheetValues.length;
    revenuesSheetValuesIndex++
  ) {
    const currentRevenueCurrency =
      revenuesSheetValues[revenuesSheetValuesIndex][
        revenuesSheetHeaders.indexOf("Currency")
      ]
    if (currentRevenueCurrency === currency) {
      const currentRevenueSubtotal =
        revenuesSheetValues[revenuesSheetValuesIndex][
          revenuesSheetHeaders.indexOf("Subtotal")
        ]
      const currentRevenueSubtotalCad =
        revenuesSheetValues[revenuesSheetValuesIndex][
          revenuesSheetHeaders.indexOf("Subtotal (CAD)")
        ]
      const currentRevenueGst =
        revenuesSheetValues[revenuesSheetValuesIndex][
          revenuesSheetHeaders.indexOf("GST")
        ]
      const currentRevenueQst =
        revenuesSheetValues[revenuesSheetValuesIndex][
          revenuesSheetHeaders.indexOf("QST")
        ]
      if (currentRevenueSubtotal !== "") {
        subtotal += currentRevenueSubtotal
      }
      if (currentRevenueSubtotal !== "" && currentRevenueSubtotalCad === "") {
        const range = revenuesSheet.getRange(
          revenuesSheetValuesIndex + 1,
          revenuesSheetHeaders.indexOf("Subtotal (CAD)") + 1
        )
        revenuesSheet.setActiveSelection(range)
        SpreadsheetApp.flush()
        throw new Error(`Missing data at ${range.getA1Notation()}`)
      }
      if (currentRevenueSubtotalCad !== "") {
        subtotalCad += currentRevenueSubtotalCad
      }
      if (currentRevenueGst !== "") {
        gst += currentRevenueGst
      }
      if (currentRevenueQst !== "") {
        qst += currentRevenueQst
      }
    }
  }
  revenueReportSheet
    .getRange("A2:D2")
    .setValues([[subtotal, subtotalCad, gst, qst]])
  revenueReportSheet.getDataRange().setFontFamily("Roboto Mono")
  revenueReportSheet
    .getRange("A2:A")
    .setNumberFormat(`0.${"0".repeat(decimalPlace)}`)
  revenueReportSheet
    .getRange("B2:B")
    .setNumberFormat(`0.${"0".repeat(decimalPlace)}`)
  revenueReportSheet
    .getRange("C2:C")
    .setNumberFormat(`0.${"0".repeat(decimalPlace)}`)
  revenueReportSheet
    .getRange("D2:D")
    .setNumberFormat(`0.${"0".repeat(decimalPlace)}`)
  DriveApp.getFileById(revenueReportSheet.getId()).moveTo(folder)
}

const generateReports = () => {
  const currenciesSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Currencies")
  const currenciesSheetValues = currenciesSheet
    .getRange(
      1,
      1,
      currenciesSheet.getLastRow(),
      currenciesSheet.getLastColumn()
    )
    .getValues()
  const currenciesSheetHeaders = currenciesSheetValues[0]
  for (
    let currenciesSheetValuesIndex = 1;
    currenciesSheetValuesIndex < currenciesSheetValues.length;
    currenciesSheetValuesIndex++
  ) {
    const currency =
      currenciesSheetValues[currenciesSheetValuesIndex][
        currenciesSheetHeaders.indexOf("Name")
      ]
    const decimalPlace =
      currenciesSheetValues[currenciesSheetValuesIndex][
        currenciesSheetHeaders.indexOf("Decimal place")
      ]
    generateExpenseReport(currency, decimalPlace)
    generateRevenueReport(currency, decimalPlace)
  }
}

const onEdit = (event: GoogleAppsScript.Events.SheetsOnEdit) => {
  const row = event.range.getRow()
  const column = event.range.getColumn()
  const value = event.range.getValue()
  const sheet = SpreadsheetApp.getActiveSheet()
  const sheetName = sheet.getName()
  const sheetValues = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues()
  const sheetHeaders = sheetValues[0]
  if (
    sheetName === "Expenses" &&
    column === sheetHeaders.indexOf("Supplier") + 1
  ) {
    const categoryRange = sheet.getRange(
      row,
      sheetHeaders.indexOf("Category") + 1
    )
    const currencyRange = sheet.getRange(
      row,
      sheetHeaders.indexOf("Currency") + 1
    )
    if (value !== "") {
      const suppliersSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Suppliers")
      const suppliersSheetValues = suppliersSheet
        .getRange(
          1,
          1,
          suppliersSheet.getLastRow(),
          suppliersSheet.getLastColumn()
        )
        .getValues()
      const suppliersSheetHeaders = suppliersSheetValues[0]
      for (
        let suppliersSheetValuesIndex = 1;
        suppliersSheetValuesIndex < suppliersSheetValues.length;
        suppliersSheetValuesIndex++
      ) {
        const name =
          suppliersSheetValues[suppliersSheetValuesIndex][
            suppliersSheetHeaders.indexOf("Name")
          ]
        const defaultExpenseCategory =
          suppliersSheetValues[suppliersSheetValuesIndex][
            suppliersSheetHeaders.indexOf("Default expense category")
          ]
        const defaultCurrency =
          suppliersSheetValues[suppliersSheetValuesIndex][
            suppliersSheetHeaders.indexOf("Default currency")
          ]
        if (name === value) {
          categoryRange.setValue(defaultExpenseCategory)
          currencyRange.setValue(defaultCurrency)
          break
        }
      }
    } else {
      categoryRange.clearContent()
      currencyRange.clearContent()
    }
  } else if (
    sheetName === "Expenses" &&
    column === sheetHeaders.indexOf("Subtotal") + 1
  ) {
    const currency = sheetValues[row - 1][sheetHeaders.indexOf("Currency")]
    if (currency === "CAD") {
      const range = sheet.getRange(
        row,
        sheetHeaders.indexOf("Subtotal (CAD)") + 1
      )
      range.setValue(value)
      if (value !== "") {
        sheet
          .getRange(row, sheetHeaders.indexOf("GST") + 1, 1, 2)
          .setFormulas([
            [
              `=${range.getA1Notation()}*Taxes!B2`,
              `=${range.getA1Notation()}*Taxes!B3`,
            ],
          ])
      } else {
        sheet
          .getRange(row, sheetHeaders.indexOf("GST") + 1, 1, 2)
          .clearContent()
      }
    }
  } else if (
    sheetName === "Expenses" &&
    column === sheetHeaders.indexOf("Subtotal (CAD)") + 1
  ) {
    const supplier = sheetValues[row - 1][sheetHeaders.indexOf("Supplier")]
    if (event.range.getValue() !== "") {
      const suppliersSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Suppliers")
      const suppliersSheetValues = suppliersSheet
        .getRange(
          1,
          1,
          suppliersSheet.getLastRow(),
          suppliersSheet.getLastColumn()
        )
        .getValues()
      const suppliersSheetHeaders = suppliersSheetValues[0]
      for (
        let suppliersSheetValuesIndex = 1;
        suppliersSheetValuesIndex < suppliersSheetValues.length;
        suppliersSheetValuesIndex++
      ) {
        const name =
          suppliersSheetValues[suppliersSheetValuesIndex][
            suppliersSheetHeaders.indexOf("Name")
          ]
        const taxable =
          suppliersSheetValues[suppliersSheetValuesIndex][
            suppliersSheetHeaders.indexOf("Taxable")
          ]
        if (name === supplier && taxable === "No") {
          return
        }
      }
      sheet
        .getRange(row, sheetHeaders.indexOf("GST") + 1, 1, 2)
        .setFormulas([
          [
            `=${event.range.getA1Notation()}*Taxes!B2`,
            `=${event.range.getA1Notation()}*Taxes!B3`,
          ],
        ])
    } else {
      sheet.getRange(row, sheetHeaders.indexOf("GST") + 1, 1, 2).clearContent()
    }
  } else if (
    sheetName === "Revenues" &&
    column === sheetHeaders.indexOf("Subtotal") + 1
  ) {
    const currency = sheetValues[row - 1][sheetHeaders.indexOf("Currency")]
    if (currency === "CAD") {
      const range = sheet.getRange(
        row,
        sheetHeaders.indexOf("Subtotal (CAD)") + 1
      )
      range.setValue(value)
      if (value !== "") {
        sheet
          .getRange(row, sheetHeaders.indexOf("GST") + 1, 1, 2)
          .setFormulas([
            [
              `=${range.getA1Notation()}*Taxes!B2`,
              `=${range.getA1Notation()}*Taxes!B3`,
            ],
          ])
      } else {
        sheet
          .getRange(row, sheetHeaders.indexOf("GST") + 1, 1, 2)
          .clearContent()
      }
    }
  } else if (
    sheetName === "Revenues" &&
    column === sheetHeaders.indexOf("Subtotal (CAD)") + 1
  ) {
    if (value !== "") {
      sheet
        .getRange(row, sheetHeaders.indexOf("GST") + 1, 1, 2)
        .setFormulas([
          [
            `=${event.range.getA1Notation()}*Taxes!B2`,
            `=${event.range.getA1Notation()}*Taxes!B3`,
          ],
        ])
    } else {
      sheet.getRange(row, sheetHeaders.indexOf("GST") + 1, 1, 2).clearContent()
    }
  }
}
