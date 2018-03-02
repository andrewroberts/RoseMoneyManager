// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 23 September 2015 06:10 GMT+1
/* jshint asi: true */
/* jshint loopfunc: true */

/*
 * Copyright (C) 2015 Andrew Roberts
 * 
 * This program is free software: you can redistribute it and/or modify it under
 * the terms of the GNU General Public License as published by the Free Software
 * Foundation, either version 3 of the License, or (at your option) any later 
 * version.
 * 
 * This program is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License along with 
 * this program. If not, see http://www.gnu.org/licenses/.
 */

/**
 *
 */

function onEdit(event) {

  Log.functionEntryPoint()

// Logger.log(event) // {range=Range, source=Spreadsheet, value=Amazong, authMode=FULL}

  var range = event.range
  var row = range.getRow()
  var column = range.getColumn()
  var sheet = SpreadsheetApp.getActiveSheet()
  var sheetName = sheet.getName()
  
  if (sheetName !== TRANSACTION_SHEET_NAME) {
    return
  }
  
  var rowRange = sheet.getRange(row, 1, 1, sheet.getLastColumn())
  
  if (column === TRANSACTION_SHEET_AMOUNT_COLUMN_NUMBER) {
  
    // If a new event is added to the amounts colon then check if it is 
    // an expense and if so make it negative

    var category = sheet
      .getRange(row, TRANSACTION_SHEET_CATEGORY_COLUMN_NUMBER)
      .getValue()
    
    if ((category.indexOf("Expenses") !== -1 ||
        category.indexOf("Liabilities") !== -1) &&
        event.value > 0) {
    
      range.setValue(event.value * -1)
    }
  }
  
} // onEdit()

/**
 * 'on open' add a new transaction to the bottom of the sheet
 */

function onOpen() {

  var ui = SpreadsheetApp.getUi()
  
  ui
    .createMenu('Rose Money Manager')
    .addItem('Add current transaction',           'addCurrentTransaction')
    .addItem('Add visa transaction',              'addCreditCardTransaction') 
    .addItem('Remove transaction',                'removeTransaction') 
    .addSeparator()
    .addItem('Sort transactions by date',         'sortTransactions') 
    .addItem('Reconcile transactions',            'reconcile')     
    .addItem('Mark transactions as Reconciled',   'markTransactionsReconciled') 
    .addSeparator()
    .addItem('Get the next 40 days transactions', 'convertCalendarEventsToExpenses')
    .addToUi()
    
  var response = ui.alert('Would you like to create a new transaction?', ui.ButtonSet.YES_NO)

  if (response === ui.Button.YES) {
    addCreditCardTransaction()  
  }  
  
} // onOpen

/**
 * Look for matching transactions and reconcile them
 */
 
function reconcile(txns) {

  Log.init({
    level: LOG_LEVEL, 
    sheetId: SpreadsheetApp.getActive().getId(),
    displayFunctionNames: LOG_DISPLAY_FUNCTION_NAMES})

  Log.functionEntryPoint()

  var txns = SpreadsheetApp
      .getActive()
      .getSheetByName(TRANSACTION_SHEET_NAME)
      .getDataRange()
      .getValues()
      
  txns.shift()    

  var timeZone = Session.getScriptTimeZone()
  
  var thisMonth = 1 // Feb
  var accountName = 'Assets:Andrew Barclays Current'
    
  var foundTxn = getNextTxn()
  
  if (foundTxn === null) {
    SpreadsheetApp.getUi().alert('Not found a txn that needs reconciling.')
    return
  }
  
  var matchingTxn = getMatchingTxn()
  
  if (matchingTxn === null) {
    SpreadsheetApp.getUi().alert('Can not find a matching txn for row ' + foundTxn.rowNumber + '.')
    return
  }
  
  // Display the two options in a dialog
  
  var template = HtmlService.createTemplateFromFile('Reconcile')
  
  template.foundRowNumber = foundTxn.rowNumber
  template.matchingRowNumber = matchingTxn.rowNumber

  template.foundTxnHtml = 
    '<td>' + foundTxn.rowNumber + '</td>' +   
    '<td>' + foundTxn.date + '</td>' + 
    '<td>' + foundTxn.description + '</td>' + 
    '<td>' + foundTxn.amount + '</td>' +
    '<td>' + foundTxn.reconciled + '</td>'
                                                               
  template.matchingTxnHtml = 
    '<td>' + matchingTxn.rowNumber + '</td>' +     
    '<td>' + matchingTxn.date + '</td>' + 
    '<td>' + matchingTxn.description + '</td>' +
    '<td>' + matchingTxn.amount + '</td>' + 
    '<td>' + matchingTxn.reconciled + '</td>'

  var html = template.evaluate().setHeight(135).setWidth(600)
    
  SpreadsheetApp.getUi().showModelessDialog(html, 'Reconcile transactions...')

  return
  
  // Private Functions
  // -----------------

  /**
   * Search for a non-reconciled txn - 'N' in this month
   */

  function getNextTxn() {

    var foundTxn = null
    
    for (var rowIndex = 0; rowIndex < txns.length; rowIndex++) {
    
      var txn = txns[rowIndex]      
      var nextDate        = txn[TRANSACTION_SHEET_DATE_COLUMN_NUMBER - 1]
      var nextMonth       = nextDate.getMonth()
      var nextAccountName = txn[TRANSACTION_SHEET_ACCOUNT_NAME_COLUMN_NUMBER - 1]
      var nextReconciled  = txn[TRANSACTION_SHEET_RECONCILED_COLUMN_NUMBER - 1]
    
      if (nextReconciled === 'N' && nextMonth === thisMonth && nextAccountName === accountName) {
      
        foundTxn = {
          month       : nextMonth,
          accountName : nextAccountName,
          reconciled  : nextReconciled,
          amount      : txn[TRANSACTION_SHEET_AMOUNT_COLUMN_NUMBER - 1],
          description : txn[TRANSACTION_SHEET_DESCRIPTION_COLUMN_NUMBER - 1],
          date        : Utilities.formatDate(nextDate, timeZone, 'YYYY-MM-dd'),
          reconciled  : nextReconciled,
          rowNumber   : rowIndex + 2, // 0-index and header
        }
      }
    }
  
    if (foundTxn === null) {
      Log.info('Found no txn that needs reconciling')
      return null
    }
    
    Log.fine('found txn: ' + JSON.stringify(foundTxn))

    return foundTxn

  } // reconcile.getNextTxn()

  /**
   * Search for a matching txn with no value in 'R' column
   */
  
  function getMatchingTxn() {
  
    var matchingTxn = null 
    
    for (var rowIndex = 0; rowIndex < txns.length; rowIndex++) {
    
      var txn = txns[rowIndex]
    
      var nextDate        = txn[TRANSACTION_SHEET_DATE_COLUMN_NUMBER - 1]
      var nextMonth       = nextDate.getMonth()
      var nextAccountName = txn[TRANSACTION_SHEET_ACCOUNT_NAME_COLUMN_NUMBER - 1]
      var nextReconciled  = txn[TRANSACTION_SHEET_RECONCILED_COLUMN_NUMBER - 1]
      var nextAmount      = txn[TRANSACTION_SHEET_AMOUNT_COLUMN_NUMBER - 1]
    
      if (nextReconciled  === '' &&
          nextMonth       === foundTxn.month && 
          nextAccountName === foundTxn.accountName &&
          nextAmount      === foundTxn.amount) {
      
        matchingTxn = {
          month       : nextMonth,
          accountName : nextAccountName,
          reconciled  : nextReconciled,
          amount      : nextAmount,
          description : txn[TRANSACTION_SHEET_DESCRIPTION_COLUMN_NUMBER - 1],
          date        : Utilities.formatDate(nextDate, timeZone, 'YYYY-MM-dd'),
          reconciled  : 'EMPTY',   
          rowNumber   : rowIndex + 2, // 0-index and header          
        }
      }
    }
    
    if (matchingTxn === null) {
      Logg.info('Found no matching txn')
      return null
    }  
  
    Log.info('found matching txn: ' + JSON.stringify(matchingTxn))
    return matchingTxn

  } // reconcile.getMatchingTxn()

} // reconcile() 

function reconcile2Txns(rowNumbers) {

  Log.init({
    level: LOG_LEVEL, 
    sheetId: SpreadsheetApp.getActive().getId(),
    displayFunctionNames: LOG_DISPLAY_FUNCTION_NAMES
  })

  Log.functionEntryPoint()

  Log.info(JSON.stringify(rowNumbers))
  
  if (rowNumbers.hasOwnProperty('found') && 
      rowNumbers.hasOwnProperty('matching') && 
      typeof rowNumbers.found !== 'undefined' &&
      typeof rowNumbers.matching !== 'undefined') {
  
    var txnsSheet = SpreadsheetApp.getActive().getSheetByName(TRANSACTION_SHEET_NAME)
    
    txnsSheet.getRange(rowNumbers.found, TRANSACTION_SHEET_RECONCILED_COLUMN_NUMBER).setValue('C')
    txnsSheet.getRange(rowNumbers.matching, TRANSACTION_SHEET_RECONCILED_COLUMN_NUMBER).setValue('D')  
    
  } else {
  
    Log.fine('Bad row numbers')
  }

  reconcile()
  
} // reconcileTxns()

/**
 * Reconcile the selected transactions
 */
 
function markTransactionsReconciled() {

  var sheet = SpreadsheetApp.getActiveSheet()

  if (sheet.getName() !== TRANSACTION_SHEET_NAME) {
    return
  }

  var wholeRange = SpreadsheetApp.getActiveRange()
  var values = wholeRange.getValues()
  var firstRow = wholeRange.getRow()
  var reconciled = []
  
  values.forEach(function() {
    reconciled.push(['R'])
  })
  
  var reconciledColumnRange = sheet
    .getRange(
      firstRow, 
      TRANSACTION_SHEET_RECONCILED_COLUMN_NUMBER, 
      values.length,
      1)
    .setValues(reconciled)
  
  return
 
} // markTransactionsReconciled

/** 
 *
 */

function sortTransactions() {

  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(TRANSACTION_SHEET_NAME)
    .sort(1)

} // sortTransactions()

/** 
 *
 */

function addCurrentTransaction() {

  addTransaction({
    name: ANDREW_BARCLAYS_CURRENT_NAME,
  })

} // addCreditCardTransaction()

/** 
 *
 */

function addCreditCardTransaction() {

  addTransaction({
    name: VISA_NAME,
  })

} // addCreditCardTransaction()

/** 
 *
 */

function addTransaction(transaction) {
  
  var sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(TRANSACTION_SHEET_NAME)
      
  var date = transaction.date || new Date()
  date = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/YYYY')
  var name = transaction.name || VISA_NAME
  var number = transaction.number || ''
  var description = transaction.description || ''
  var memo = transaction.memo || ''
  var category = transaction.category || ''
  var reconciled = transaction.reconciled || 'N'
  var amount = transaction.amount || 0
  
  sheet
    .appendRow([
      date,
      name,
      number,
      description,
      memo,
      category,
      reconciled,
      amount])
    .getRange(sheet.getLastRow(), TRANSACTION_SHEET_DESCRIPTION_COLUMN_NUMBER)
    .activate()

} // addTransaction()

/**
 * Remove the transactions in the active row
 */

function removeTransaction() {

  var sheet = SpreadsheetApp.getActiveSheet() 
  var rowNumber = sheet.getActiveCell().getRow()
  sheet.deleteRow(rowNumber)
    
} // removeTransaction()

/**
 * Import Visa CSV/GSheet
 */
 
function importVisa() {

  SpreadsheetApp
    .openById(VISA_CSV_SHEET_ID)
    .getSheets()[0]
    .getDataRange()
    .getValues()
    .forEach(function(row) { 
      addTransaction({
        date: row[0],
        name: VISA_NAME,
        description: row[1],
        memo: row[4],
        amount: row[6] * -1,
    })
  })

  return

} // importVisa()
  
function test() {
  var value = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'YYYY/MM/dd')
  var a = SpreadsheetApp.getActiveSheet().getActiveRange().setValue(value)
  return
}

/*  
  // Find a slot before the future transactions
  // ------------------------------------------

  if (lastRow === 1) {
  
    transactionsSheet.getRange(2, 1).setValue(new Date())
    return
  }

  // Find the slot for the new row (earlier than any future transactions

  // TODO - This assumes they are in date order

  var dates = transactionsSheet
    .getRange(2, 1, lastRow)
    .getValues()
    
  // Remove the header  
  dates.shift()  
    
  var today = (new Date()).getTime()
  var rowNumberToAddNew = lastRow
    
  var result = dates.some(function(row) {
  
    if (rowNumberToAddNew === 1) {
    
      Logger.log('Back to header')
      return true
    }
  
    if (row[0].getTime() > today) {
    
      Logger.log('Decrement rowNumberToAddNew')
      rowNumberToAddNew--
      
    } else {
    
      Logger.log('')
      return true
    }
  })
*/