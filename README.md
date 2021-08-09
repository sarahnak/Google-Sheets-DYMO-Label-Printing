# Google-Sheets-DYMO-Label-Printing
Google Apps Script project to print Google Sheets form data to a DYMO LabelWriter printer. Used [this repo](https://github.com/errontitus/dymo_cryo_label) and [the DYMO Developer SDK Support Blog](https://developers.dymo.com/2010/06/02/dymo-label-framework-javascript-library-samples-print-a-label/).

## To use:

1. Install [the DYMO Label Software (Windows)](https://s3.amazonaws.com/download.dymo.com/dymo/Software/Win/DLS8Setup8.7.4.exe).
Other versions are available [here](https://www.labelvalue.com/dymo-software-and-drivers).

2. Copy and paste [Code.gs](Code.gs), [MySidebarFunctions.gs](MySidebarFunctions.gs), and [Sidebar remote DYMO.html](https://github.com/sarahnak/Google-Sheets-DYMO-Label-Printing/blob/d380ec3c022a1033c735e0d0290567d09dfca11a/Sidebar%20remote%20DYMO.html) into a new Google Apps Script project attached to a spreadsheet. Name the first sheet "Registration".

3. Connect a DYMO LabelWriter printer. 

4. From the Google Sheets menu bar, go to "Labels"-> "Print Labels". In the likely case that your spreadsheet isn't configured exactly like the one I attached this project to, you'll get an error message. The preview should show "Sample address text here." 

5. Click print. 

## Customizing

### Label text

Currently, the project takes the last line of the spreadsheet and formats it as a JavaScript string literal (getFormattedLast in Code.gs). This is set to the text value of the label in [Sidebar remote DYMO.html](https://github.com/sarahnak/Google-Sheets-DYMO-Label-Printing/blob/d380ec3c022a1033c735e0d0290567d09dfca11a/Sidebar%20remote%20DYMO.html). 

You can just write your own method to replace getFormattedLast.

### Label XML
This just uses the default address label XML provided in [the sample by DYMO](https://developers.dymo.com/2010/06/02/dymo-label-framework-javascript-library-samples-print-a-label/), but if you wanted to change it you could design your own layout in the Dymo Label Software and copy/paste the text of the .label file into the labelXml variable.

### Printers used
I use the first LabelWriter printer that's connected. You can change how the printerNameFinal variable in the Sidebar remote DYMO.html file is set to something that works for you.
