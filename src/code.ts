const scriptProperties = PropertiesService.getScriptProperties()

interface ColumnIds {
  [name: string]: number
}

const getColumnIds = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
  let columnIds: ColumnIds = {}
  for (let headerIndex = 0; headerIndex < headers.length; headerIndex++) {
    columnIds[headers[headerIndex]] = headerIndex + 1
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

const getAbsoluteA1Notation = (range: GoogleAppsScript.Spreadsheet.Range) => {
  const parts = range.getA1Notation().match(/([A-Z]+)(\d+)/)
  return `$${parts[1]}$${parts[2]}`
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
    {
      name: "Show sheetsdrive sidebar",
      functionName: "showSheetsdriveSidebar",
    },
    { name: "Update exchange rates", functionName: "updateExchangeRates" },
  ]
  sheet.addMenu("Custom utilities", menuEntries)
}

const showSheetsdriveSidebar = () => {
  const template = HtmlService.createTemplateFromFile("upload")
  template.webAppUrl = scriptProperties.getProperty("webAppUrl")
  const html = template
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle("Sheetsdrive")
  SpreadsheetApp.getUi().showSidebar(html)
}

const generateExpenseReport = () => {
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
    `${sheetFilename} expense report`
  )
  expenseReportSheet
    .getRange("A1:G1")
    .setFontWeight("bold")
    .setValues([
      [
        "GIFI",
        "Category",
        "Subtotal (CAD)",
        "GST",
        "QST",
        "Percentage used for business activities",
        "Capital expense",
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
  let expenseReportSheetRow = 2
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
    const expenseCategoryGifi =
      expenseCategoriesSheetValues[expenseCategoriesSheetValuesIndex][
        expenseCategoriesSheetHeaders.indexOf("GIFI")
      ]
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
      if (currentExpenseCategory === expenseCategoryName) {
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
        if (currentExpenseSubtotalCad === "") {
          DriveApp.getFileById(expenseReportSheet.getId()).setTrashed(true)
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
    if (subtotalCad !== 0 || gst !== 0 || qst !== 0) {
      expenseReportSheet
        .getRange(`A${expenseReportSheetRow}:G${expenseReportSheetRow}`)
        .setValues([
          [
            expenseCategoryGifi,
            expenseCategoryName,
            subtotalCad,
            gst,
            qst,
            expenseCategoryPercentageUsedForBusinessActivities,
            expenseCategoryCapitalExpense,
          ],
        ])
      expenseReportSheetRow++
    }
  }
  expenseReportSheet.getDataRange().setFontFamily("Roboto Mono")
  expenseReportSheet.getRange("A2:A").setNumberFormat("0")
  expenseReportSheet.getRange("B2:B").setNumberFormat("@")
  expenseReportSheet.getRange("C2:C").setNumberFormat("#,##0.00")
  expenseReportSheet.getRange("D2:D").setNumberFormat("#,##0.00")
  expenseReportSheet.getRange("E2:E").setNumberFormat("#,##0.00")
  expenseReportSheet.getRange("F2:F").setNumberFormat("0.00%")
  expenseReportSheet.getRange("G2:G").setNumberFormat("@")
  DriveApp.getFileById(expenseReportSheet.getId()).moveTo(folder)
}

const generateRevenueReport = () => {
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
    `${sheetFilename} revenue report`
  )
  revenueReportSheet
    .getRange("A1:E1")
    .setFontWeight("bold")
    .setValues([["GIFI", "Category", "Subtotal (CAD)", "GST", "QST"]])
  const revenueCategoriesSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Revenue categories")
  const revenueCategoriesSheetValues = revenueCategoriesSheet
    .getRange(
      1,
      1,
      revenueCategoriesSheet.getLastRow(),
      revenueCategoriesSheet.getLastColumn()
    )
    .getValues()
  const revenueCategoriesSheetHeaders = revenueCategoriesSheetValues[0]
  let revenueReportSheetRow = 2
  for (
    let revenueCategoriesSheetValuesIndex = 1;
    revenueCategoriesSheetValuesIndex < revenueCategoriesSheetValues.length;
    revenueCategoriesSheetValuesIndex++
  ) {
    const revenueCategoryName =
      revenueCategoriesSheetValues[revenueCategoriesSheetValuesIndex][
        revenueCategoriesSheetHeaders.indexOf("Name")
      ]
    const revenueCategoryGifi =
      revenueCategoriesSheetValues[revenueCategoriesSheetValuesIndex][
        revenueCategoriesSheetHeaders.indexOf("GIFI")
      ]
    let subtotalCad = 0
    let gst = 0
    let qst = 0
    for (
      let revenuesSheetValuesIndex = 1;
      revenuesSheetValuesIndex < revenuesSheetValues.length;
      revenuesSheetValuesIndex++
    ) {
      const currentRevenueCategory =
        revenuesSheetValues[revenuesSheetValuesIndex][
          revenuesSheetHeaders.indexOf("Category")
        ]
      if (currentRevenueCategory === revenueCategoryName) {
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
        const currentRevenueRecurrence =
          revenuesSheetValues[revenuesSheetValuesIndex][
            revenuesSheetHeaders.indexOf("Recurrence")
          ]
        if (currentRevenueSubtotalCad === "") {
          DriveApp.getFileById(revenueReportSheet.getId()).setTrashed(true)
          const range = revenuesSheet.getRange(
            revenuesSheetValuesIndex + 1,
            revenuesSheetHeaders.indexOf("Subtotal (CAD)") + 1
          )
          revenuesSheet.setActiveSelection(range)
          SpreadsheetApp.flush()
          throw new Error(`Missing data at ${range.getA1Notation()}`)
        }
        if (currentRevenueSubtotalCad !== "") {
          subtotalCad +=
            currentRevenueRecurrence !== ""
              ? currentRevenueSubtotalCad * currentRevenueRecurrence
              : currentRevenueSubtotalCad
        }
        if (currentRevenueGst !== "") {
          gst +=
            currentRevenueRecurrence !== ""
              ? currentRevenueGst * currentRevenueRecurrence
              : currentRevenueGst
        }
        if (currentRevenueQst !== "") {
          qst +=
            currentRevenueRecurrence !== ""
              ? currentRevenueQst * currentRevenueRecurrence
              : currentRevenueQst
        }
      }
    }
    if (subtotalCad !== 0 || gst !== 0 || qst !== 0) {
      revenueReportSheet
        .getRange(`A${revenueReportSheetRow}:E${revenueReportSheetRow}`)
        .setValues([
          [revenueCategoryGifi, revenueCategoryName, subtotalCad, gst, qst],
        ])
      revenueReportSheetRow++
    }
  }
  revenueReportSheet.getDataRange().setFontFamily("Roboto Mono")
  revenueReportSheet.getRange("A2:A").setNumberFormat("0")
  revenueReportSheet.getRange("B2:B").setNumberFormat("@")
  revenueReportSheet.getRange("C2:C").setNumberFormat("#,##0.00")
  revenueReportSheet.getRange("D2:D").setNumberFormat("#,##0.00")
  revenueReportSheet.getRange("E2:E").setNumberFormat("#,##0.00")
  DriveApp.getFileById(revenueReportSheet.getId()).moveTo(folder)
}

const generateReports = () => {
  generateExpenseReport()
  generateRevenueReport()
}

interface Rates {
  [date: string]: number
}

const getPreviousRate = (startDate: Date, date: Date, rates: Rates) => {
  const currentDate = new Date(date)
  while (currentDate >= startDate) {
    const rate = rates[currentDate.toISOString().split("T")[0]]
    if (rate) {
      return rate
    }
    currentDate.setDate(currentDate.getDate() - 1)
  }
}

const getNextRate = (endDate: Date, date: Date, rates: Rates) => {
  const currentDate = new Date(date)
  while (currentDate <= endDate) {
    const rate = rates[currentDate.toISOString().split("T")[0]]
    if (rate) {
      return rate
    }
    currentDate.setDate(currentDate.getDate() + 1)
  }
}

interface BtcToCadData {
  data: {
    quotes: [
      {
        timeOpen: string
        quote: {
          high: number
          low: number
        }
      }
    ]
  }
}

const btcToCad = (from: string, to: string) => {
  const timeStart =
    Math.floor(new Date(`${from}T00:00:00.000Z`).getTime() / 1000) - 1
  const timeEnd =
    Math.ceil(new Date(`${to}T23:59:59.999Z`).getTime() / 1000) + 1
  const response = UrlFetchApp.fetch(
    `https://api.coinmarketcap.com/data-api/v3.1/cryptocurrency/historical?id=1&timeStart=${timeStart}&timeEnd=${timeEnd}&interval=1d&convertId=2784&format=json`
  )
  const json: BtcToCadData = JSON.parse(response.getContentText())
  const rates: Rates = {}
  for (const quote of json.data.quotes) {
    rates[quote.timeOpen.split("T")[0]] =
      Math.round(((quote.quote.high + quote.quote.low) / 2) * 100) / 100
  }
  return rates
}

interface UsdToCadData {
  observations: [
    {
      d: string
      FXUSDCAD: {
        v: string
      }
    }
  ]
}

const usdToCad = (from: string, to: string) => {
  const extendedFromDate = new Date(from)
  extendedFromDate.setDate(extendedFromDate.getDate() - 7)
  const startDate = extendedFromDate.toISOString().split("T")[0]
  const extendedToDate = new Date(to)
  extendedToDate.setDate(extendedToDate.getDate() + 7)
  const endDate = extendedToDate.toISOString().split("T")[0]
  const response = UrlFetchApp.fetch(
    `https://www.bankofcanada.ca/valet/observations/FXUSDCAD/json?start_date=${startDate}&end_date=${endDate}`
  )
  const json: UsdToCadData = JSON.parse(response.getContentText())
  const extendedRates: Rates = {}
  for (const observation of json.observations) {
    extendedRates[observation.d] = parseFloat(observation.FXUSDCAD.v)
  }
  const rates: Rates = {}
  const currentDate = new Date(from)
  while (currentDate <= new Date(to)) {
    const formattedCurrentDate = currentDate.toISOString().split("T")[0]
    if (!extendedRates[formattedCurrentDate]) {
      const previousRate = getPreviousRate(
        extendedFromDate,
        currentDate,
        extendedRates
      )
      const nextRate = getNextRate(extendedToDate, currentDate, extendedRates)
      if (previousRate && nextRate) {
        rates[formattedCurrentDate] =
          Math.round(((previousRate + nextRate) / 2) * 10000) / 10000
      }
    } else {
      rates[formattedCurrentDate] = extendedRates[formattedCurrentDate]
    }
    currentDate.setDate(currentDate.getDate() + 1)
  }
  return rates
}

const updateExchangeRates = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
  const reportingPeriodSheet = sheet.getSheetByName("Reporting period")
  const from = Utilities.formatDate(
    reportingPeriodSheet.getRange("A2").getValue(),
    "America/Montreal",
    "yyyy-MM-dd"
  )
  const to = Utilities.formatDate(
    reportingPeriodSheet.getRange("B2").getValue(),
    "America/Montreal",
    "yyyy-MM-dd"
  )
  const values = []
  const btcToCadRates = btcToCad(from, to)
  const usdToCadRates = usdToCad(from, to)
  const currentDate = new Date(from)
  while (currentDate <= new Date(to)) {
    const formattedCurrentDate = currentDate.toISOString().split("T")[0]
    const btcToCadRate = btcToCadRates[formattedCurrentDate]
    const usdToCadRate = usdToCadRates[formattedCurrentDate]
    if (btcToCadRate && usdToCadRate) {
      values.push([formattedCurrentDate, btcToCadRate, usdToCadRate])
    }
    currentDate.setDate(currentDate.getDate() + 1)
  }
  const exchangeRatesSheet = sheet.getSheetByName("Exchange rates")
  exchangeRatesSheet.getDataRange().clearContent()
  exchangeRatesSheet.getRange("A1:C1").setValues([["Date", "BTC", "USD"]])
  exchangeRatesSheet.getRange(`A2:C${values.length + 1}`).setValues(values)
}

const setExpenseTaxValues = (
  row: number,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  sheetHeaders: string[],
  supplier: string
) => {
  const suppliersSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Suppliers")
  const suppliersSheetValues = suppliersSheet
    .getRange(1, 1, suppliersSheet.getLastRow(), suppliersSheet.getLastColumn())
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
    if (name === supplier) {
      if (taxable === "Yes") {
        break
      } else if (taxable === "No") {
        return
      }
    }
  }
  const subtotalCadA1 = sheet
    .getRange(row, sheetHeaders.indexOf("Subtotal (CAD)") + 1)
    .getA1Notation()
  const absoluteGstHeaderA1 = getAbsoluteA1Notation(
    sheet.getRange(1, sheetHeaders.indexOf("GST") + 1)
  )
  const absoluteQstHeaderA1 = getAbsoluteA1Notation(
    sheet.getRange(1, sheetHeaders.indexOf("QST") + 1)
  )
  sheet
    .getRange(row, sheetHeaders.indexOf("GST") + 1, 1, 2)
    .setFormulas([
      [
        `=${subtotalCadA1}*VLOOKUP(${absoluteGstHeaderA1}, Taxes!A2:B1000, 2, FALSE)`,
        `=${subtotalCadA1}*VLOOKUP(${absoluteQstHeaderA1}, Taxes!A2:B1000, 2, FALSE)`,
      ],
    ])
}

const setRevenueTaxValues = (
  row: number,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  sheetHeaders: string[],
  subtotalCadA1: string
) => {
  const absoluteGstHeaderA1 = getAbsoluteA1Notation(
    sheet.getRange(1, sheetHeaders.indexOf("GST") + 1)
  )
  const absoluteQstHeaderA1 = getAbsoluteA1Notation(
    sheet.getRange(1, sheetHeaders.indexOf("QST") + 1)
  )
  sheet
    .getRange(row, sheetHeaders.indexOf("GST") + 1, 1, 2)
    .setFormulas([
      [
        `=${subtotalCadA1}*VLOOKUP(${absoluteGstHeaderA1}, Taxes!A2:B1000, 2, FALSE)`,
        `=${subtotalCadA1}*VLOOKUP(${absoluteQstHeaderA1}, Taxes!A2:B1000, 2, FALSE)`,
      ],
    ])
}

const onEdit = (event: GoogleAppsScript.Events.SheetsOnEdit) => {
  if (event.range.getNumRows() !== 1 || event.range.getNumColumns() !== 1) {
    return
  }
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
    const supplier = sheetValues[row - 1][sheetHeaders.indexOf("Supplier")]
    const currency = sheetValues[row - 1][sheetHeaders.indexOf("Currency")]
    const dateRange = sheet.getRange(row, sheetHeaders.indexOf("Date") + 1)
    const subtotalRange = sheet.getRange(
      row,
      sheetHeaders.indexOf("Subtotal") + 1
    )
    const subtotalCadRange = sheet.getRange(
      row,
      sheetHeaders.indexOf("Subtotal (CAD)") + 1
    )
    const exchangeRatesSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Exchange rates")
    const exchangeRatesSheetValues = exchangeRatesSheet
      .getRange(
        1,
        1,
        exchangeRatesSheet.getLastRow(),
        exchangeRatesSheet.getLastColumn()
      )
      .getValues()
    const exchangeRatesSheetHeaders = exchangeRatesSheetValues[0]
    if (value !== "") {
      if (subtotalRange.getFormula() === "") {
        if (currency === "CAD") {
          subtotalCadRange.setValue(value)
        } else {
          subtotalCadRange.setFormula(
            `=${event.range.getA1Notation()}*VLOOKUP(${dateRange.getA1Notation()}, 'Exchange rates'!A2:C1000, ${
              exchangeRatesSheetHeaders.indexOf(currency) + 1
            }, FALSE)`
          )
        }
        setExpenseTaxValues(row, sheet, sheetHeaders, supplier)
      }
    } else {
      sheet
        .getRange(row, sheetHeaders.indexOf("Subtotal (CAD)") + 1, 1, 3)
        .clearContent()
    }
  } else if (
    sheetName === "Expenses" &&
    column === sheetHeaders.indexOf("Subtotal (CAD)") + 1
  ) {
    const supplier = sheetValues[row - 1][sheetHeaders.indexOf("Supplier")]
    if (event.range.getValue() !== "") {
      setExpenseTaxValues(row, sheet, sheetHeaders, supplier)
    } else {
      sheet.getRange(row, sheetHeaders.indexOf("GST") + 1, 1, 2).clearContent()
    }
  } else if (
    sheetName === "Revenues" &&
    column === sheetHeaders.indexOf("Subtotal") + 1
  ) {
    const currency = sheetValues[row - 1][sheetHeaders.indexOf("Currency")]
    const dateRange = sheet.getRange(row, sheetHeaders.indexOf("Date") + 1)
    const subtotalRange = sheet.getRange(
      row,
      sheetHeaders.indexOf("Subtotal") + 1
    )
    const subtotalCadRange = sheet.getRange(
      row,
      sheetHeaders.indexOf("Subtotal (CAD)") + 1
    )
    const exchangeRatesSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Exchange rates")
    const exchangeRatesSheetValues = exchangeRatesSheet
      .getRange(
        1,
        1,
        exchangeRatesSheet.getLastRow(),
        exchangeRatesSheet.getLastColumn()
      )
      .getValues()
    const exchangeRatesSheetHeaders = exchangeRatesSheetValues[0]
    if (value !== "") {
      if (subtotalRange.getFormula() === "") {
        if (currency === "CAD") {
          subtotalCadRange.setValue(value)
        } else {
          subtotalCadRange.setFormula(
            `=${event.range.getA1Notation()}*VLOOKUP(${dateRange.getA1Notation()}, 'Exchange rates'!A2:C1000, ${
              exchangeRatesSheetHeaders.indexOf(currency) + 1
            }, FALSE)`
          )
        }
      }
    } else {
      subtotalCadRange.clearContent()
    }
  } else if (
    sheetName === "Revenues" &&
    column === sheetHeaders.indexOf("Taxable") + 1
  ) {
    const subtotalCadRange = sheet.getRange(
      row,
      sheetHeaders.indexOf("Subtotal (CAD)") + 1
    )
    if (value === "Yes") {
      setRevenueTaxValues(
        row,
        sheet,
        sheetHeaders,
        subtotalCadRange.getA1Notation()
      )
    } else {
      sheet.getRange(row, sheetHeaders.indexOf("GST") + 1, 1, 2).clearContent()
    }
  } else if (
    sheetName === "Shareholders" &&
    column === sheetHeaders.indexOf("Deposit amount") + 1
  ) {
    const depositCurrency =
      sheetValues[row - 1][sheetHeaders.indexOf("Deposit currency")]
    const currentValueCadRange = sheet.getRange(
      row,
      sheetHeaders.indexOf("Current value (CAD)") + 1
    )
    const exchangeRatesSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Exchange rates")
    const exchangeRatesSheetValues = exchangeRatesSheet
      .getRange(
        1,
        1,
        exchangeRatesSheet.getLastRow(),
        exchangeRatesSheet.getLastColumn()
      )
      .getValues()
    const exchangeRatesSheetHeaders = exchangeRatesSheetValues[0]
    if (value !== "") {
      if (depositCurrency === "CAD") {
        currentValueCadRange.setValue(value)
      } else {
        currentValueCadRange.setFormula(
          `=${event.range.getA1Notation()}*VLOOKUP('Reporting period'!$B$2, 'Exchange rates'!A2:C1000, ${
            exchangeRatesSheetHeaders.indexOf(depositCurrency) + 1
          }, FALSE)`
        )
      }
    } else {
      currentValueCadRange.clearContent()
    }
  } else if (
    sheetName === "Shareholder loans" &&
    column === sheetHeaders.indexOf("Amount") + 1
  ) {
    const currency = sheetValues[row - 1][sheetHeaders.indexOf("Currency")]
    const dateRange = sheet.getRange(row, sheetHeaders.indexOf("Date") + 1)
    const amountCadRange = sheet.getRange(
      row,
      sheetHeaders.indexOf("Amount (CAD)") + 1
    )
    const exchangeRatesSheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Exchange rates")
    const exchangeRatesSheetValues = exchangeRatesSheet
      .getRange(
        1,
        1,
        exchangeRatesSheet.getLastRow(),
        exchangeRatesSheet.getLastColumn()
      )
      .getValues()
    const exchangeRatesSheetHeaders = exchangeRatesSheetValues[0]
    if (value !== "") {
      if (currency === "CAD") {
        amountCadRange.setValue(value)
      } else {
        amountCadRange.setFormula(
          `=${event.range.getA1Notation()}*VLOOKUP(${dateRange.getA1Notation()}, 'Exchange rates'!A2:C1000, ${
            exchangeRatesSheetHeaders.indexOf(currency) + 1
          }, FALSE)`
        )
      }
    } else {
      amountCadRange.clearContent()
    }
  }
}
