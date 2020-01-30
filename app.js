const excel = require('exceljs');
const fs = require('fs');
const express = require('express');
const app = express();
const cors = require('cors');
const bodyParser = require("body-parser");

const corsoptions = {
    origin: '*',
    methods: 'GET,POST,PUT,DELETE',
    allowedHeaders: 'Content-Type,Response-Type',
    optionsSuccessStatus: 200
};

app.use(cors(corsoptions));

app.use(express.json());

app.use(bodyParser.urlencoded({extended: true}));

var examdatajson = fs.readFileSync('./json/examdata.json');
var examdata = JSON.parse(examdatajson);
var exams = examdata.exams;

function getexams(data, exams) {
    var outdata = [];

    var datalen = Object.keys(data).length;

    for (num = 0; num < datalen; num++) {
        var exam = data[num].exam;
        var examdata = exams[exam];
        outdata.push(examdata);
    };

    return outdata;
};

async function editxlsx(data) {
    var book = new excel.Workbook();
    book = await book.xlsx.readFile('./xlsx/table.xlsx');
    var sheet = book.getWorksheet('data');

    var datalen = Object.keys(data).length;

    for (num = 0; num < datalen; num++) {
        var exam = data[num]
        var row = exam.indexR;
        var row2 = row + 1;
        var row3 = row + 2;
        var cell = exam.indexC;

        for (i = 0; i < 3; i++) {

            sheet.getRow(row).getCell(cell).value = exam.cell1;
            sheet.getRow(row2).getCell(cell).value = exam.cell2;
            sheet.getRow(row3).getCell(cell).value = exam.cell3;

            sheet.getRow(row).getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFD966' },
                bgColor: { argb: 'FFFFD966' }
            };

            sheet.getRow(row2).getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFD966' },
                bgColor: { argb: 'FFFFD966' }
            };

            sheet.getRow(row3).getCell(cell).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFD966' },
                bgColor: { argb: 'FFFFD966' }
            };

        };

    };

    await book.xlsx.writeFile('./xlsx/temp.xlsx');
    var file = fs.readFileSync('./xlsx/temp.xlsx');
    return file;
};

app.post('/xlsx', async (req, res) => { 

    var intdata = req.body.exams;
    
    var data = getexams(intdata, exams);

    var file = await editxlsx(data);

    res.send(file.toString('base64'));

});

app.listen(process.env.PORT || 3000, function () {
    console.log('[ Port: %d ][ Mode: %s ]', this.address().port, app.settings.env);
});