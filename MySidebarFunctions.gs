//Will read from the JS library hosted on my Github account. Works
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('DYMO Blood Sample Labels')
  SpreadsheetApp.getUi().showSidebar(html);
};