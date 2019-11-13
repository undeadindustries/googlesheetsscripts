# googlesheetsscripts
Useful scripts for Google sheets

Each of these is coded to be used by itself. If you know javascript, you'll sort out how to combine them if needed.

These are also built for my specific needs. You will probably have to edit them to do exactly what you need them to.

Simply copy and paste code into your Google Sheet's Code.gs. I gave these files .js filenames so my IDE would format them correctly.
Obviously, you get to Code.js by clicking tools->Script Editor.

sortSheets.js- 
Orders sheets/tabs in a Google sheet by name.

summaryPage.js- 
If you have many sheets/tabs in a Google sheet, you can create a "Summary" sheet. This script will populate all of column A to a horizontal row in the Summary sheet. This way, you have one summary you can reference. It helps to lock the Summary sheet so no one edits it on purpose. You must edit the sheets it references. This also assumes you have a tab called "template" that you can copy from to create new tabs. It skips that tab.

Javascript isn't my favorite language. I'm sure this can be coded more elegantly... I welcome suggestions, but please be nice.
