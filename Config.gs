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

var SCRIPT_NAME = "RoseMoneyManager"
var SCRIPT_VERSION = "v1.0"

var VISA_NAME = 'Liabilities:Andrew Bcard Visa'
var ANDREW_BARCLAYS_CURRENT_NAME = 'Assets:Andrew Barclays Current'

// Expenses Calendar
// -----------------

var CALENDAR_NAME = 'Regular Transactions'

// Logging
// -------

var LOG_LEVEL                  = Log.Level.ALL
// var LOG_SHEET_ID               = TRANSACTION_SHEET_ID
var LOG_DISPLAY_FUNCTION_NAMES = Log.DisplayFunctionNames.NO

// Transaction Sheet
// -----------------

var TRANSACTION_SHEET_NAME = 'Txns'

var ReconciledColours = Object.freeze({
  'N': '#f4cccc', // light red 3
  'C': '#fff2cc', // light yellow 3
  'R': '#d9ead3', // light green 3
  'D': '#d9d2e9', // light purple 3		
})

var TRANSACTION_SHEET_NOT_RECONCILED_COLOUR = '#f4cccc'

var TRANSACTION_SHEET_DATE_COLUMN_NUMBER         = 1
var TRANSACTION_SHEET_ACCOUNT_NAME_COLUMN_NUMBER = 2
var TRANSACTION_SHEET_NUMBER_COLUMN_NUMBER       = 3
var TRANSACTION_SHEET_DESCRIPTION_COLUMN_NUMBER  = 4
var TRANSACTION_SHEET_MEMO_COLUMN_NUMBER         = 5
var TRANSACTION_SHEET_CATEGORY_COLUMN_NUMBER     = 6
var TRANSACTION_SHEET_RECONCILED_COLUMN_NUMBER   = 7
var TRANSACTION_SHEET_AMOUNT_COLUMN_NUMBER       = 8
