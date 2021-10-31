/* node Stocks.js --url = "https://www.equitymaster.com/stockquotes/1-69/list-of-nifty-100-stocks"--dest = "stocksWala.html"
--source = stocksWala.html --destFolder=StockFolder --csv=Stocks.csv
 */
// node Stocks.js 
let minimist = require("minimist");
let fs = require("fs");

//to download html from web it will provide us browsers functionality
let axios = require("axios");
let jsdom = require("jsdom");
let pdflib = require("pdf-lib");
let Excel = require("excel4node");
const { convert } = require("html-to-text");
const { excelColor } = require("excel4node/distribution/lib/types");
const { red } = require("excel4node/distribution/lib/types/excelColor");
const { rgb, createPDFAcroField, PDFDocument } = require("pdf-lib");
let path = require('path');

let clargs = minimist(process.argv);

//to download html from web axios will provide us browsers functionality
let downloadKaPromise = axios.get(clargs.url);
downloadKaPromise.then(function(response) {
    let html = response.data;
    fs.writeFileSync(clargs.dest, html, "utf-8");
})


dom manupulation using jsdom

fs.readFile(clargs.source, "utf-8", function(err, html) {
    let dom = new jsdom.JSDOM(html);
    let document = dom.window.document;

    let allStocks = document.querySelectorAll("tr[style]");


    let allStocksfifty = [];
    for (let i = 0; i < allStocks.length; i++) {

        let singleStock = allStocks[i];

        let individualStock = {
            name: '',
            BSEPrice: '',
            NSEPrice: '',
            Note: " ***ALL PRICE(RS) and %(UP OR DOWN)***"
        }


        let stockName = singleStock.querySelector("td.alignleft > a");

        individualStock.name = convert(stockName.textContent);

        let Price = singleStock.querySelectorAll("td.alignright");

        individualStock.BSEPrice = convert(Price[0].textContent);
        individualStock.NSEPrice = convert(Price[1].textContent);



        allStocksfifty.push(individualStock);


    }
    // now creat the json
    let stockKiJson = JSON.stringify(allStocksfifty);

    fs.writeFileSync(clargs.json, stockKiJson, 'utf-8');
});




now creat the Excel
let stockJson = fs.readFileSync("Stocks.json", 'utf-8');
let JsoStock = JSON.parse(stockJson);
let wb = new Excel.Workbook();

let myStyle = wb.createStyle({
    font: {
        bold: true,
        color: '#0000FF',
        underline: true,
        size: 20,

    },

    alignment: {
        horizontal: ['center']
    },
});


let myStyle2 = wb.createStyle({
    alignment: {
        horizontal: 'center',
    },
});

let red_s = wb.createStyle({
    font: {
        color: "#FF0000",
    },

    alignment: {
        horizontal: 'center'
    }
});

let green_s = wb.createStyle({
    font: {
        color: "#00b04c",
    },

    alignment: {
        horizontal: 'center'
    }
})

for (let i = 0; i < JsoStock.length; i++) {
    let sheet = wb.addWorksheet(JsoStock[i].name);

    //cell(start row,start col,end row,end col, merge)
    sheet.cell(1, 9, 1, 10, true).string("Stock Name").style(myStyle);
    sheet.cell(2, 9, 2, 10, true).string(JsoStock[i].name).style(myStyle2);

    sheet.cell(5, 9, 5, 10, true).string("BSE Prices").style(myStyle);

    JsoStock[i].BSEPrice = JsoStock[i].BSEPrice.replace(' ', '    ');
    if (JsoStock[i].BSEPrice.includes('-')) {
        sheet.cell(6, 9, 6, 10, true).string(JsoStock[i].BSEPrice).style(red_s);
    } else {
        sheet.cell(6, 9, 6, 10, true).string(JsoStock[i].BSEPrice).style(green_s);
    }

    sheet.cell(9, 9, 9, 10, true).string("NSE Prices").style(myStyle);

    JsoStock[i].NSEPrice = JsoStock[i].NSEPrice.replace(' ', '    ');
    if (JsoStock[i].NSEPrice.includes('-')) {
        sheet.cell(10, 9, 10, 10, true).string(JsoStock[i].NSEPrice).style(red_s);
    } else {
        sheet.cell(10, 9, 10, 10, true).string(JsoStock[i].NSEPrice).style(green_s);
    }

    sheet.cell(13, 7, 13, 12, true).string(JsoStock[i].Note).style(myStyle);

}

wb.write(clargs.csv);


//first make folders using path
fs.mkdirSync("StockFolder");
let JsonfileR = fs.readFileSync("Stocks.json", 'utf-8');
let JsoStock = JSON.parse(JsonfileR);

for (let i = 0; i < JsoStock.length; i++) {
    let folderName = path.join(clargs.destFolder, JsoStock[i].name);
    fs.mkdirSync(folderName);

    let folderInsideName = path.join(clargs.destFolder, JsoStock[i].name, JsoStock[i].name + ".pdf");
    createPDF(JsoStock[i].name, JsoStock[i].BSEPrice, JsoStock[i].NSEPrice, folderInsideName);
}

// write the made pdf
function createPDF(nameS, Bse, Nse, folderInsideName) {

    let name = nameS;
    let BseP = Bse;
    let NseP = Nse;




    let OrgpdfBytes = fs.readFileSync("stocks.pdf");
    let pdfBytesloadKiPromises = pdflib.PDFDocument.load(OrgpdfBytes);

    pdfBytesloadKiPromises.then(function(pdfDoc) {

        let page = pdfDoc.getPage(0);
        page.drawText(name, {
            x: 320,
            y: 945,
            size: 25,
            bold: true
        });

        if (Bse.includes('-')) {
            page.drawText(BseP, {
                x: 330,
                y: 698,
                size: 25,
                color: rgb(1, 0, 0),
            })
        } else {
            page.drawText(BseP, {
                x: 330,
                y: 698,
                size: 25,
                color: rgb(0, 0.69, 0.27)

            })
        }

        if (Nse.includes('-')) {
            page.drawText(NseP, {
                x: 330,
                y: 470,
                size: 25,
                color: rgb(1, 0, 0),
            })
        } else {
            page.drawText(NseP, {
                x: 330,
                y: 470,
                size: 25,
                color: rgb(0, 0.69, 0.27),


            })
        }

        let saveKiPromise = pdfDoc.save();
        saveKiPromise.then(function(newpdfBytes) {
            fs.writeFileSync(folderInsideName, newpdfBytes);
        })
    })
}
