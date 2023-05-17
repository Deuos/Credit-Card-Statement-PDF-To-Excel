import { PdfReader } from "pdfreader";
import XLSX from 'xlsx'
import fs from 'fs'

//Path to Pdf/Output
const pdfPath = 'Add pdf path';
const excelPath = 'excel/output.xlsx';
const exclusionListPath = 'exclusionList.txt'

const dataForDateOrigin = 'A1'
const dataForBalanceOrigin = 'A5'
const dataForInfoOrigin = 'A10'


//Writes to the xlsx
function writeToXlsx(dataForDate, dataForBalance, dataForInfo) {

    if (typeof XLSX == 'undefined') XLSX = require('xlsx');
    
    const workbook = XLSX.utils.book_new();
    //Name of Sheet
    const sheetName = 'Sheet1';

    const worksheet = XLSX.utils.json_to_sheet([]);

    // Populate dataForDate starting from cell A1
    XLSX.utils.sheet_add_json(worksheet, dataForDate, { origin: dataForDateOrigin });

    // Populate dataForBalance starting from cell A10
    XLSX.utils.sheet_add_json(worksheet, dataForBalance, { origin: dataForBalanceOrigin });

    // Populate dataForInfo starting from cell A20
    XLSX.utils.sheet_add_json(worksheet, dataForInfo, { origin: dataForInfoOrigin });

    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

    XLSX.writeFile(workbook, excelPath);
}

function convertDataToReadable(dataForDate, dataForBalance, dataForInfo) {

    //Open Date - Close Date
    const finalDataForDate = []
    //Total Points Earned Month - Total Points - Current Balance
    const finalDataForBalance = []
    //Post Date - Trans Date - Ref # - Description - Amount
    const finalDataForInfo = []

    /* FinalDataForDate */
    // Format the "From" date
    const fromDate = dataForDate.slice(0, 6).flat().join('');
    finalDataForDate.push({ From: fromDate });
    // Format the "To" date
    const toDate = dataForDate.slice(6).flat().join('');
    finalDataForDate.push({ To: toDate });
    /* FinalDataForDate */
    /* FinalDataForBalance */
    const pointsEarned = dataForBalance[0][0]
    finalDataForBalance.push({ PointsEarned: pointsEarned })
    const totalPointsEarned = dataForBalance[2][0]
    finalDataForBalance.push({ TotalPointsEarned: totalPointsEarned })
    const balance = dataForBalance[4][0]
    finalDataForBalance.push({ Balance: balance })
    /* FinalDataForBalance */
    /* FinalDataForInfo */

    //Just a extra incase, the scraper doesn't work correctly
    if (dataForInfo[0][0] === '2022') {

        dataForInfo.splice(0, 1);
    }
    for (let i = 0; i < dataForInfo.length; i += 9) {

        if (
            dataForInfo[i] &&
            dataForInfo[i + 1] &&
            dataForInfo[i + 2] &&
            dataForInfo[i + 3] &&
            dataForInfo[i + 4] &&
            dataForInfo[i + 5] &&
            dataForInfo[i + 6] &&
            dataForInfo[i + 7] &&
            dataForInfo[i + 8]
        ) {
            const entry = {
                'PostDate': `${dataForInfo[i][0]}/${dataForInfo[i + 2][0]}`,
                'TransDate': `${dataForInfo[i + 3][0]}/${dataForInfo[i + 5][0]}`,
                'Reference': dataForInfo[i + 6][0],
                'Description': dataForInfo[i + 7][0],
                'Amount': dataForInfo[i + 8][0]
            };

            finalDataForInfo.push(entry);
        }
    }
    /* FinalDataForInfo */

    console.log(finalDataForDate);
    console.log(finalDataForBalance);
    console.log(finalDataForInfo);

    writeToXlsx(finalDataForDate, finalDataForBalance, finalDataForInfo, excelPath);
}

/* Main Reads Pdf and Calls to write to XLSX*/
function convertPDFToExcel(pdfPath) {

    //Pairs to Include in the xlsx
    const keywordPairs = [
        //Enter start and where to stop
        //{ start: "ExampleStart", stop: "ExampleStop" },
    ]

    //Pairs to Exclude from the xlsx
    const excludedPairs = [
        //{ start: "ExampleStart", stop: "ExampleStop" },
        // Add more excluded pairs as needed
    ];

    //List links to exclusionList and enter words or phrases to exclude
    let exclusionList = [];

    //read exclusion file and then convert
    fs.readFile(exclusionListPath, 'utf8', function (err, data) {

        //Console.log error
        if (err) {
            console.error('Error reading exclusion list:', err);
            return;
        }

        exclusionList = data.split('\n').map(item => item.trim()).filter(item => item !== '');
        //console.log('Exclusion list:', exclusionList);

        //PDFreader indicators
        let withinRange = false;
        let currentPairIndex = 0;
        let excludeLogging = false;
        //Accumulated Data
        const dataForDate = [];
        const dataForBalance = [];
        const dataForInfo = [];

        new PdfReader().parseFileItems(pdfPath, function (err, item) {

            //Error
            if (err) {
                //err
                console.error("error:", err);
            }
            //When end of the file
            else if (!item) {
                convertDataToReadable(dataForDate, dataForBalance, dataForInfo)
                console.warn("end of file");
            }
            //Console.logs and adds to the Xlsx
            else if (item.text) {
                if (currentPairIndex < keywordPairs.length) {
                    const currentPair = keywordPairs[currentPairIndex];
                    const startFound = currentPair && item.text.includes(currentPair.start);
                    const stopFound = currentPair && item.text.includes(currentPair.stop);

                    if (startFound && !withinRange && !excludeLogging) {
                        withinRange = true;
                        console.log("Start logging from keyword:", currentPair.start);
                    } else if (stopFound && withinRange && !excludeLogging) {
                        withinRange = false;
                        console.log("Stop logging at keyword:", currentPair.stop);
                        // Perform any required action when the stop keyword is found.

                        // Move to the next keyword pair
                        currentPairIndex++;
                        if (currentPairIndex >= keywordPairs.length) {
                            console.log("All keyword pairs processed.");
                            // Perform any required action when all keyword pairs have been processed.
                        }
                    }

                    const excludedPairStartFound = excludedPairs.some(pair => item.text.includes(pair.start));
                    const excludedPairStopFound = excludedPairs.some(pair => item.text.includes(pair.stop));

                    if (excludedPairStartFound && !excludeLogging) {
                        excludeLogging = true;
                        console.log("Start excluding within pair:", item.text);
                        return; // Skip logging the item
                    } else if (excludeLogging && excludedPairStopFound) {
                        excludeLogging = false;
                        console.log("Stop excluding within pair:", item.text);
                        // Perform any required action when the end of the excluded pair is found.
                    }

                    const excludeWords = exclusionList.some(word => item.text.includes(word));

                    if (withinRange && !excludeWords && !excludeLogging) {
                        if (currentPair.start === "Open Date:") {
                            dataForDate.push([item.text])
                        }
                        else if (currentPair.start === "New Balance" || currentPair.start === "Earned This Statement") {
                            dataForBalance.push([item.text])
                        }
                        else if (currentPair.start === "Transactions") {
                            dataForInfo.push([item.text])
                        }
                        //console.log(item.text);
                        // Perform any action or logic within the desired range, excluding words from the exclusion list.
                    }
                }
            }

            // console.log("DataForDate--------------------")
            // console.log(dataForDate)
            // console.log("DataForBalance--------------------")
            // console.log(dataForBalance)
            //console.log("DataForInfo--------------------")
            //console.log(dataForInfo)

        });

    });
}

function readFullPdf(pdfPath){

    new PdfReader().parseFileItems(pdfPath, (err, item) => {
                if (err) console.error("error:", err);
                else if (!item) console.warn("end of file");
                else if (item.text) console.log(item.text);
            });
}


convertPDFToExcel(pdfPath);
//readFullPdf(pdfPath)