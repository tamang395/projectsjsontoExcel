const express = require('express');
const path = require('path')
const isJson = require('is-json');
const fs = require('fs');
const xl = require('excel4node');
const bodyParser = require('body-parser');
const hbs = require('hbs');
const { type } = require('os');
const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb' }));
var wb = new xl.Workbook();
var ws = wb.addWorksheet('data');
const viewsPathDirectory = path.join(__dirname, 'views')// this provides a location of vews where this index.hbs//it will directly target vies folder
app.set('view engine', 'hbs')//this tell which template engine is used
app.set('views', viewsPathDirectory);//accessing viwe path 


// app.use(express.static(viewsPathDirectory))//serve up the static files using the reference of pulic folder of index. html
app.get('', (req, res) => {
    res.render('index', {
        title: "ExcelApi",

    })
})
app.post('/jsontoexcel', (req, res) => {
    const jsondata = req.body.json;
    const object = JSON.parse(jsondata);
    const legs1Data = object.data[0].legs;
    // console.log(legs1Data);
    var arry = [];
    for (var obj of legs1Data) {
        var rowdata = {};
        rowdata['Trade No'] = ""
        rowdata['lots'] = obj.lots;
        rowdata['legName'] = obj.legName
        rowdata['entryDate'] = obj.entryDate;
        rowdata['strikePrice'] = obj.strikePrice;
        rowdata['buyOrSell'] = obj.buyOrSell;
        rowdata['futuresOrOptions'] = obj.futuresOrOptions;
        rowdata['entryValue'] = obj.entryValue;
        rowdata['exitDate'] = obj.exitDate;
        rowdata['exitValue'] = obj.exitValue;
        // rowdata['equity'] = obj.equity;//accessing the value against the key
        rowdata['profits'] = (obj.exitValue - obj.entryValue) * obj.lots * 75;
        const date1 = new Date(obj.entryDate);
        const date2 = new Date(obj.exitDate);
        const diffTime = Math.abs(date2 - date1);
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        rowdata['Days'] = diffDays;
        arry.push(rowdata)
    }
    var arrayData = ['Trade No', 'lots', 'legName', 'entryDate', 'strikePrice', 'buyOrSell',
        'futuresOrOptions', 'entryValue', 'exitDate', 'exitValue', 'Days', 'profits'];//put require column
    var i = 1;
    var style = wb.createStyle({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '2172d7',

        }
    });
    var style2 = wb.createStyle({
        fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: 'FF8080',

        }
    });
    for (var timeP of arrayData) {

        ws.cell(1, i).string(timeP)
        i++;
    }

    for (var i = 0; i < arry.length; i++) {
        for (var j = 0; j < arrayData.length; j++) {
            if (typeof (arry[i][arrayData[j]]) === "string") {
                ws.cell(i + 2, j + 1).string(arry[i][arrayData[j]])

            } else {
                ws.cell(i + 2, j + 1).number(arry[i][arrayData[j]])
                ws.cell(7, 13).formula('L2+L3+L4+L5+L6+L7')
                ws.cell(1, 13).string('Total Profit')
                ws.cell(2, 1).number(1)

            }
        }
        ws.cell(2, 12, 7).style(style2)
        ws.cell(2, 13, 7).style(style2)



        ws.cell(2, 7, 7).style(style);
        ws.cell(2, 9, 7).style(style);
        ws.cell(2, 10, 7).style(style);


    }
    wb.write("Excel.xlsx", res)
    console.log(arry);
})
app.listen(3000, () => {
    console.log("App listting port 3000");
})