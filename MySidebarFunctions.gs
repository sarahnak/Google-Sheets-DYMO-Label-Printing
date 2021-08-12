//Will read from the JS library hosted on my Github account. Works
function showSidebarRemoteDYMO() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar remote DYMO')
      .setTitle('DYMO Blood Sample Labels')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
};