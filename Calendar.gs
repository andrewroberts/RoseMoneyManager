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
 * Convert calendar events to expenses.
 * 
 * Read in all of todays events and add a row to the transaction list 
 * spreadsheet for each.
 */
 
function convertCalendarEventsToExpenses() {

  Log.init({
    level: LOG_LEVEL, 
    sheetId: LOG_SHEET_ID,
    displayFunctionNames: LOG_DISPLAY_FUNCTION_NAMES})
    
  Log.functionEntryPoint()
  
  // Check the calendars
  // -------------------
  
  var calendars = CalendarApp.getCalendarsByName(CALENDAR_NAME)
  
  if (calendars.length === 0) {
    
    throw new Error('There is no calendar called ' + 
                      CALENDAR_NAME + '.')
  }
  
  if (calendars.length > 1) {
    
    throw new Error('Found ' + calendars.length + ' calendars ' + 
                      'called ' + CALENDAR_NAME + ' ' + 
                      'when there should only be one.')
  }
  
  // Check today's events
  // --------------------
  //
  // The events are stored in UTC so we need to allow for that by 
  // defining todays date in terms of UTC to avoid picking up events 
  // either side of today.
  
  var today = new Date()
  
  var startOfDay = today
  
  startOfDay.setUTCHours(0)
  startOfDay.setMinutes(0)
  startOfDay.setSeconds(0)
  startOfDay.setMilliseconds(0)  
  
  // Check the next 40 days ahead
  var endOfDay = new Date(startOfDay.getTime() + 40 * 24 * 60 * 60 * 1000)
  
  var events = calendars[0].getEvents(startOfDay, endOfDay)
  
  if (events.length === 0) {
    
    Log.info('No events today')
    return
  } 
  
  Log.info('Number of events today: ' + events.length + 
             ' start: ' + startOfDay + 
             ' end: ' + endOfDay)
  
  // Store these events as expenses
  // ------------------------------
  
  for (var eventIndex = 0; eventIndex < events.length; eventIndex++) {
    var event = events[eventIndex]
    var details = event.getDescription()
    var json = JSON.parse(details)
    json.date = event.getAllDayEndDate()
    addTransaction(json)  
  }
    
} // convertCalendarEventsToExpenses()

// Test functions
// --------------

function test_json() {

  var object = {
    name: "Assets:Andrew Barclays Current", 
    description:"Swalec", 
    category: "Expenses:Bills:Electricity", 
    amount: "23"
  }
  
  var string = JSON.stringify(object)
  
//  var string = '{name: "Assets:Andrew Barclays Current", description:"Swalec", category: "Expenses:Bills:Electricity", amount: "23"}'
  return
}


