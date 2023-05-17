
# Credit-Card-Statement-PDF-To-Excel

Converts Pdf into Excel

Made for Elan Financial Services

### Installation

    *npm install*

    *node server.js*

### Converting to Pdf

Add pdfs to*pdf* folder

First run readFullPdf

Find all start and stops, from what keyword to start and what keyword to stop, and include them into keywordPairs, additionally, you can include the start and stop from within the keywordPairs using excludedPairs.

Additionally, you can include single phrases in exclusionList to eliminate them

Finally, run convertPDFToExcel and it will take tinkering for it to work depending on your pdf.
