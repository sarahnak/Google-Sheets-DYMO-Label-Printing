// Since there were issues with the site connection being marked as insecure when the script ran, only resolved by commenting out the import statement, here are 3 of my attempts to fix it. 


//Will read from the JS library hosted on my Github account. Works
function showSidebarRemoteDYMO() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar remote DYMO')
      .setTitle('DYMO Blood Sample Labels')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
};

//Imports the JS library as a .html file in this project. Works but not used
function showSidebarLocalDYMO() {
  var htmlOutput = HtmlService.createTemplateFromFile('Sidebar local DYMO')
      .evaluate()
      .setTitle('DYMO Blood Sample Labels (Local JS file)')
      // .getCode();
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);

};

// Loading the HTMLtemplate for 'Sidebar local DYMO.html' with this XML code in it wouldn't compile - this is the other workaround. I think it's something with how the server reads/transforms the HTML.
//Loading as a template is necessary for the sidebar using the 'DYMO Label Framework 3.0.html' file, since scriptlets (ie the import statement) are only executed through being HTMLtemplates
function getLabelXml(){
  return '<?xml version="1.0" encoding="utf-8"?>\
    <DieCutLabel Version="8.0" Units="twips">\
        <PaperOrientation>Landscape</PaperOrientation>\
        <Id>Address</Id>\
        <PaperName>30252 Address</PaperName>\
        <DrawCommands/>\
        <ObjectInfo>\
            <TextObject>\
                <Name>Text</Name>\
                <ForeColor Alpha="255" Red="0" Green="0" Blue="0" />\
                <BackColor Alpha="0" Red="255" Green="255" Blue="255" />\
                <LinkedObjectName></LinkedObjectName>\
                <Rotation>Rotation0</Rotation>\
                <IsMirrored>False</IsMirrored>\
                <IsVariable>True</IsVariable>\
                <HorizontalAlignment>Left</HorizontalAlignment>\
                <VerticalAlignment>Middle</VerticalAlignment>\
                <TextFitMode>ShrinkToFit</TextFitMode>\
                <UseFullFontHeight>True</UseFullFontHeight>\
                <Verticalized>False</Verticalized>\
                <StyledText/>\
            </TextObject>\
            <Bounds X="332" Y="150" Width="4455" Height="1260" />\
        </ObjectInfo>\
    </DieCutLabel>';
}

//for the DYMO Label Framework 3.0.html in the local DYMO sidebar
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

//to debug HTMLtemplates, get the code from the console, copy into .gs file, and debug from there (from docs)
function debugTemplates() {
  Logger.log(HtmlService
      .createTemplateFromFile('Sidebar')
      .getCode());
}

