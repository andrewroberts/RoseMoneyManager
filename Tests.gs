function test() {
  var date = new Date()
  var timeZone = Session.getScriptTimeZone()
  var d = Utilities.formatDate(date, timeZone, 'YYYY-MM-dd')
  return
}
