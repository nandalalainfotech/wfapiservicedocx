const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const excel = require('exceljs');
const HTMLtoDOCX = require('html-docx-js');
const saveAs = require('file-saver');
const fs = require('fs');
var pdf = require('dynamic-html-pdf');
var html = fs.readFileSync('./template/document.html', 'utf8');

// const wellsfargoJson = fs.readFileSync('./json/wellsfargo.json', 'utf8');
// let wellsfargo1 = JSON.parse(wellsfargoJson);

var express = require('express');
var app = express();

app.use(cors());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('public'));



var PORT = process.env.PORT || 80;



app.get('/', cors(), function(req, res) {
    res.sendFile(path.join(__dirname, '/mock.html'));
});


function isEmpty(obj) {
    for (var prop in obj) {
        if (obj.hasOwnProperty(prop))
            return false;
    }
    return true;
}

app.post('/api/docx', cors(), async function(req, res) {
    console.log("calling---->docx");



    var docx = HTMLtoDOCX.asBlob(html, { orientation: 'landscape', margins: { top: 720 }, });
    saveAs(html, 'test.docx');
    fs.writeFile(html, docx, function(err) {
        res.setHeader('Content-Disposition', `attachment; filename=wellsfargo.docx`);
        res.setHeader('Content-Length', docx.length);
        res.send(docx);
    });


});


app.post('/api/pdf', cors(), function(req, res) {
    var chunks = [];
    res.on("data", function(chunk) {
        chunks.push(chunk);
    });
    res.on("end", function(chunk) {
        var body = Buffer.concat(chunks);
        // console.log(body.toString());
    });
    res.on("error", function(error) {
        // console.error(error);
    });
    // console.log("req.body----->", req.body);
    if (isEmpty(req.body)) {
        return;
    }
    let wellsfargo = req.body;
    // console.log("calling pdf req----->",wellsfargo);
    var options = {
        format: "A3",
        orientation: "landscape",
        border: "10mm",
    };

    var document = {
        type: 'buffer',
        template: html,
        context: {
            Wellsfargo: wellsfargo
        },
    };


    if (document === null) {
        return null;

    } else {
        pdf.create(document, options).then(response => {
            res.writeHead(200, {
                "Content-Disposition": "attachment;filename=" + "wellsFargo.pdf",
                'Content-Type': 'application/pdf'
            });
            return res.end(response);
        }).catch(error => {
            console.error(error)
        });
    };

});




app.post('/api/excel', cors(), function(req, res) {
    var chunks = [];
    res.on("data", function(chunk) {
        chunks.push(chunk);
    });
    res.on("end", function(chunk) {
        var body = Buffer.concat(chunks);
        // console.log(body.toString());
    });
    res.on("error", function(error) {
        // console.error(error);
    });
    if (isEmpty(req.body)) {
        return;
    }
    let wellsfargo = req.body;
    let workbook = new excel.Workbook();
    let worksheet = workbook.addWorksheet('Quote Form');

    // border none
    worksheet.views = [{ showGridLines: false }];

    worksheet.getRow(1).height = 40;
    worksheet.getRow(2).height = 30;

    worksheet.getRow(10).height = 40;
    worksheet.getRow(11).height = 20;
    worksheet.getRow(12).height = 25;
    worksheet.getRow(13).height = 25;
    worksheet.getRow(14).height = 80;

    worksheet.getRow(17).height = 170;

    worksheet.columns = [{ key: 'A', width: 8.0 }, { key: 'B', width: 10.0 }, { key: 'C', width: 18.0 },
        { key: 'D', width: 18.0 }, { key: 'E', width: 15.0 }, { key: 'F', width: 20.0 }, { key: 'G', width: 15.0 },
        { key: 'H', width: 15.0 }, { key: 'I', width: 15.0 }, { key: 'J', width: 20.0 }, { key: 'K', width: 15.0 },
        { key: 'L', width: 23.0 }, { key: 'M', width: 15.0 }, { key: 'N', width: 2.0 }, { key: 'O', width: 15.0 },
        { key: 'P', width: 15.0 }, { key: 'Q', width: 15.0 }, { key: 'R', width: 23.0 }, { key: 'S', width: 17.0 },
        { key: 'T', width: 15.0 }, { key: 'U', width: 20.0 }
    ];


    const imageId1 = workbook.addImage({
        filename: './images/wellsfargo.png',
        extension: 'png',
    });

    worksheet.addImage(imageId1, 'B1:B1', );

    worksheet.mergeCells('B1:U1');
    worksheet.getCell('B1:U1').value = wellsfargo.mainTitle;
    worksheet.getCell('B1:U1').font = {
        size: 28,
        name: 'Verdana',
        family: 1

    };
    worksheet.getCell('B1:U1').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('B1:U1').border = {
        // top: {style:'thin'},
        // left: {style:'none'},
        // bottom: {style:'thin'},
        right: { style: 'thin' }
    };
    // --------------------------COMMON-------------------------

    // worksheet.mergeCells('K3:K10');

    // ['B18:E18', 'F18', 'G18', 'H18', 'I18', 'J18', 'K18', 'L18', 'M18', 'O18', 'P18', 'Q18', 'R18', 'S18', 'T18', 'U18',
    //     'B20:E20', 'F20', 'G20', 'H20', 'I20', 'J20', 'K20', 'L20', 'M20', 'O20', 'P20', 'Q20', 'R20', 'S20', 'T20', 'U20',
    //     'B22:E22', 'F22', 'G22', 'H22', 'I22', 'J22', 'K22', 'L22', 'M22', 'O22', 'P22', 'Q22', 'R22', 'S22', 'T22', 'U22',
    //     'B24:E24', 'F24', 'G24', 'H24', 'I24', 'J24', 'K24', 'L24', 'M24', 'O24', 'P24', 'Q24', 'R24', 'S24', 'T24', 'U24',
    //     'B26:E26', 'F26', 'G26', 'H26', 'I26', 'J26', 'K26', 'L26', 'M26', 'O26', 'P26', 'Q26', 'R26', 'S26', 'T26', 'U26',
    //     'B28:E28', 'F28', 'G28', 'H28', 'I28', 'J28', 'K28', 'L28', 'M28', 'O28', 'P28', 'Q28', 'R28', 'S28', 'T28', 'U28',
    //     'B30:E30', 'F30', 'G30', 'H30', 'I30', 'J30', 'K30', 'L30', 'M30', 'O30', 'P30', 'Q30', 'R30', 'S30', 'T30', 'U30',
    //     'B32:E32', 'F32', 'G32', 'H32', 'I32', 'J32', 'K32', 'L32', 'M32', 'O32', 'P32', 'Q32', 'R32', 'S32', 'T32', 'U32',
    //     'B34:E34', 'F34', 'G34', 'H34', 'I34', 'J34', 'K34', 'L34', 'M34', 'O34', 'P34', 'Q34', 'R34', 'S34', 'T34', 'U34',
    //     'B42:G42', 'H42', 'I42', 'J42', 'K42', 'L42', 'M42', 'O42', 'P42', 'Q42', 'R42', 'S42', 'T42', 'U42', 'L43', 'U43',
    //     'B44:G44', 'H44', 'I44', 'J44', 'K44', 'L44', 'M44', 'O44', 'P44', 'Q44', 'R44', 'S44', 'T44', 'U44', 'L45', 'U45',
    //     'B46:G46', 'H46', 'I46', 'J46', 'K46', 'L46', 'M46', 'O46', 'P46', 'Q46', 'R46', 'S46', 'T46', 'U46',
    //     'T4:U4', 'T5:U5', 'T6:U6', 'T7:U7', 'T8:U8'
    // ].map(key => {
    //     worksheet.getCell(key).fill = {
    //         type: 'pattern',
    //         pattern: 'solid',
    //         fgColor: { argb: 'F2F2F2' },
    //         bgColor: { argb: 'F2F2F2' }
    //     };
    // });


    ['A1', 'A2'].map(key => {
        worksheet.getCell(key).font = {
            size: 11,
            name: 'Verdana',
            family: 1

        };
    });
    // -----------------------------------------------------------


    worksheet.mergeCells('B2:Q2');
    worksheet.getCell('B2:Q2').border = {
        top: { style: 'none' },
        left: { style: 'none' },
        bottom: { style: 'none' },
        right: { style: 'none' }
    };

    worksheet.mergeCells('R2:S2');
    worksheet.getCell('R2:S2').value = "Date Submitted";
    worksheet.getCell('R2:S2').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('R2:S2').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('R2:S2').border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('T2:U2');
    worksheet.getCell('T2:U2').value = wellsfargo.dateSubmitted;
    worksheet.getCell('T2:U2').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('T2:U2').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('T2:U2').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('T2:U2').border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('B3:J3');
    worksheet.getCell('B3:J3').value = wellsfargo.tableOneTitle;
    worksheet.getCell('B3:J3').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell('B3:J3').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('B3:J3').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('B3:J3').border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('B4:D4');
    worksheet.getCell('B4:D4').value = "Project # or Work Order #";
    worksheet.getCell('B4:D4').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('B4:D4').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('B4:D4').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('E4:J4');
    worksheet.getCell('E4:J4').value = wellsfargo.projectOrWorkOrder;
    worksheet.getCell('E4:J4').font = {
        size: 14,
        name: 'Verdana',
        family: 1
            // bold: true
    };
    worksheet.getCell('E4:J4').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('E4:J4').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('E4:J4').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };




    worksheet.mergeCells('B5:D5');
    worksheet.getCell('B5:D5').value = "WF Project/ Property Manager";
    worksheet.getCell('B5:D5').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('B5:D5').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('B5:D5').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };


    worksheet.mergeCells('E5:J5');
    worksheet.getCell('E5:J5').value = wellsfargo.wfProjectOrPropertyManager;
    worksheet.getCell('E5:J5').font = {
        size: 14,
        name: 'Verdana',
        family: 1
            // bold: true
    };
    worksheet.getCell('E5:J5').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('E5:J5').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('E5:J5').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('B6:D6');
    worksheet.getCell('B6:D6').value = "BE Number: ";
    worksheet.getCell('B6:D6').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('B6:D6').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('B6:D6').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('E6:J6');
    worksheet.getCell('E6:J6').value = wellsfargo.beNumber;
    worksheet.getCell('E6:J6').font = {
        size: 14,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('E6:J6').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('E6:J6').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('E6:J6').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('B7:D7');
    worksheet.getCell('B7:D7').value = "Building / Project Name:";
    worksheet.getCell('B7:D7').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('B7:D7').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('B7:D7').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('E7:J7');
    worksheet.getCell('E7:J7').value = wellsfargo.buildingOrProjectName;
    worksheet.getCell('E7:J7').font = {
        size: 14,
        name: 'Verdana',
        family: 1

    };
    worksheet.getCell('E7:J7').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('E7:J7').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('E7:J7').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('B8:D8');
    worksheet.getCell('B8:D8').value = "BE Service or Delivery Address";
    worksheet.getCell('B8:D8').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('B8:D8').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('B8:D8').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('E8:J8');
    worksheet.getCell('E8:J8').value = wellsfargo.beServiceOrDeliveryAddress;
    worksheet.getCell('E8:J8').font = {
        size: 14,
        name: 'Verdana',
        family: 1

    };
    worksheet.getCell('E8:J8').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('E8:J8').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };
    worksheet.getCell('E8:J8').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };

    worksheet.mergeCells('B9:D9');
    worksheet.getCell('B9:D9').value = "Project Area (sq.ft.):";
    worksheet.getCell('B9:D9').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('B9:D9').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('B9:D9').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('E9:J9');
    worksheet.getCell('E9:J9').value = wellsfargo.projectArea;
    worksheet.getCell('E9:J9').font = {
        size: 14,
        name: 'Verdana',
        family: 1,

    };
    worksheet.getCell('E9:J9').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('E9:J9').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('E9:J9').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('B10:C10');
    worksheet.getCell('B10:C10').value = "Estimated Start Date:";
    worksheet.getCell('B10:C10').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('B10:C10').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('B10:C10').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('D10:E10');
    worksheet.getCell('D10:E10').value = wellsfargo.estimatedStartDate;
    worksheet.getCell('D10:E10').font = {
        size: 14,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('D10:E10').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('D10:E10').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('D10:E10').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('F10:G10');
    worksheet.getCell('F10:G10').value = "Estimated Complete Date:";
    worksheet.getCell('F10:G10').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('F10:G10').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('F10:G10').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('H10:J10');
    worksheet.getCell('H10:J10').value = wellsfargo.estimatedCompleteDate;
    worksheet.getCell('H10:J10').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('H10:J10').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('H10:J10').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('H10:J10').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };
    //    ---------------------------2 row---------------------------

    worksheet.mergeCells('L3:U3');
    worksheet.getCell('L3:U3').value = wellsfargo.tableTwoTitle;
    worksheet.getCell('L3:U3').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell('L3:U3').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('L3:U3').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('L3:U3').border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('L4:M4');
    worksheet.getCell('L4:M4').value = "Company Name:";
    worksheet.getCell('L4:M4').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L4:M4').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('L4:M4').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('N4:Q4');
    worksheet.getCell('N4:Q4').value = wellsfargo.companyName;
    worksheet.getCell('N4:Q4').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('N4:Q4').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('N4:Q4').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('N4:Q4').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('R4:S4');
    worksheet.getCell('R4:S4').value = "WF Vendor Number:";
    worksheet.getCell('R4:S4').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('R4:S4').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('R4:S4').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T4:U4');
    worksheet.getCell('T4:U4').value = wellsfargo.wfVendOrNumber;
    worksheet.getCell('T4:U4').font = {
        size: 14,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('T4:U4').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('T4:U4').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('L5:M5');
    worksheet.getCell('L5:M5').value = "Remit To Address:";
    worksheet.getCell('L5:M5').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L5:M5').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('L5:M5').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('N5:Q5');
    worksheet.getCell('N5:Q5').value = wellsfargo.remitToAddress;
    worksheet.getCell('N5:Q5').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('N5:Q5').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('N5:Q5').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('N5:Q5').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('R5:S5');
    worksheet.getCell('R5:S5').value = "Proposal Number:";
    worksheet.getCell('R5:S5').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('R5:S5').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('R5:S5').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T5:U5');
    worksheet.getCell('T5:U5').value = wellsfargo.proposalNumber;
    worksheet.getCell('T5:U5').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('T5:U5').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('T5:U5').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('L6:M6');
    worksheet.getCell('L6:M6').value = " City, State, Zip :";
    worksheet.getCell('L6:M6').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L6:M6').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('L6:M6').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('N6:Q6');
    worksheet.getCell('N6:Q6').value = wellsfargo.cityStateZip;
    worksheet.getCell('N6:Q6').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('N6:Q6').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('N6:Q6').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('N6:Q6').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('R6:S6');
    worksheet.getCell('R6:S6').value = "WF Contract Number:";
    worksheet.getCell('R6:S6').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('R6:S6').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('R6:S6').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T6:U6');
    worksheet.getCell('T6:U6').value = wellsfargo.wfContractNumber;
    worksheet.getCell('T6:U6').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('T6:U6').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('T6:U6').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };



    worksheet.mergeCells('L7:M7');
    worksheet.getCell('L7:M7').value = "Contact Name:";
    worksheet.getCell('L7:M7').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L7:M7').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('L7:M7').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('N7:Q7');
    worksheet.getCell('N7:Q7').value = wellsfargo.contactName;
    worksheet.getCell('N7:Q7').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('N7:Q7').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('N7:Q7').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('N7:Q7').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('R7:S7');
    worksheet.getCell('R7:S7').value = "Change Order #:";
    worksheet.getCell('R7:S7').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('R7:S7').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('R7:S7').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T7:U7');
    worksheet.getCell('T7:U7').value = wellsfargo.changeOrder;
    worksheet.getCell('T7:U7').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('T7:U7').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('T7:U7').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('L8:M8');
    worksheet.getCell('L8:M8').value = "Phone";
    worksheet.getCell('L8:M8').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L8:M8').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('L8:M8').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('N8:Q8');
    worksheet.getCell('N8:Q8').value = wellsfargo.phone;
    worksheet.getCell('N8:Q8').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('N8:Q8').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('N8:Q8').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('N8:Q8').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('R8:S8');
    worksheet.getCell('R8:S8').value = "Change Order Previous PO#:";
    worksheet.getCell('R8:S8').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('R8:S8').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('R8:S8').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T8:U8');
    worksheet.getCell('T8:U8').value = wellsfargo.changeOrderPreviousPO;
    worksheet.getCell('T8:U8').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('T8:U8').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('T8:U8').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('L9:M9');
    worksheet.getCell('L9:M9').value = "Cell:";
    worksheet.getCell('L9:M9').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L9:M9').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('L9:M9').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('N9:U9');
    worksheet.getCell('N9:U9').value = wellsfargo.cell;
    worksheet.getCell('N9:U9').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('N9:U9').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('N9:U9').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('N9:U9').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('L10:M10');
    worksheet.getCell('L10:M10').value = "Email:";
    worksheet.getCell('L10:M10').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L10:M10').alignment = { vertical: 'middle', horizontal: 'right' };
    worksheet.getCell('L10:M10').border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('N10:U10');
    worksheet.getCell('N10:U10').value = wellsfargo.email;
    worksheet.getCell('N10:U10').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('N10:U10').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('N10:U10').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('N10:U10').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    //    ------------------------------------------------------

    // ------------------------------3 row--------------------------
    worksheet.mergeCells('B11:U11');

    // ------------------------------4 row--------------------------

    worksheet.mergeCells('B12:U12');
    worksheet.getCell('B12:U12').value = wellsfargo.scopeTitle;
    worksheet.getCell('B12:U12').font = {
        size: 16,
        name: 'Verdana',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell('B12:U12').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('B12:U12').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('B12:U12').border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('B13:U13');
    worksheet.getCell('B13:U13').value = wellsfargo.scopeSubHeadOne;
    worksheet.getCell('B13:U13').font = {
        size: 11,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('B13:U13').alignment = { vertical: 'top', horizontal: 'center', wrapText: true };
    worksheet.getCell('B13:U13').border = {
        top: { style: 'none' },
        left: { style: 'thick' },
        bottom: { style: 'none' },
        right: { style: 'thick' }
    };



    worksheet.mergeCells('B14:U14');
    worksheet.getCell('B14:U14').value = {
        'richText': [

            { 'font': { 'size': 14, 'name': 'Verdana', 'family': 1 }, 'text': wellsfargo.scopeSubDescription_D1 },

            { 'font': { 'size': 14, 'name': 'Verdana', 'family': 1, 'color': { 'argb': 'FF0000' } }, 'text': wellsfargo.scopeSubDescription_D2 }

        ]
    };
    // worksheet.getCell('B14:U14').value = wellsfargo.scopeSubDescription;
    worksheet.getCell('B14:U14').font = {
        size: 14,
        name: 'Verdana',
        family: 1,
        // bold: true
    };
    worksheet.getCell('B14:U14').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'C5D9F1' },
        bgColor: { argb: 'C5D9F1' }
    };
    worksheet.getCell('B14:U14').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('B14:U14').border = {
        top: { style: 'none' },
        left: { style: 'thick' },
        bottom: { style: 'none' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('B15:U15');
    worksheet.getCell('B15:U15').value = wellsfargo.installHeading;
    worksheet.getCell('B15:U15').font = {
        size: 16,
        name: 'Verdana',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell('B15:U15').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('B15:U15').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('B15:U15').border = {
        top: { style: 'none' },
        left: { style: 'thick' },
        bottom: { style: 'none' },
        right: { style: 'thick' }
    };

    worksheet.getRow(16).height = 25;
    worksheet.mergeCells('B16:E16');
    worksheet.getCell('B16:E16').value = wellsfargo.detailSubHeadOne;
    worksheet.getCell('B16:E16').font = {
        size: 16,
        name: 'Verdana',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell('B16:E16').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('B16:E16').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('B16:E16').border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('F16:M16');
    worksheet.getCell('F16:M16').value = wellsfargo.productSubHeadTwo;
    worksheet.getCell('F16:M16').font = {
        size: 16,
        name: 'Verdana',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell('F16:M16').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('F16:M16').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('F16:M16').border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('O16:T16');
    worksheet.getCell('O16:T16').value = wellsfargo.laborSubHeadThree;
    worksheet.getCell('O16:T16').font = {
        size: 16,
        name: 'Verdana',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell('O16:T16').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('O16:T16').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('O16:T16').border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };



    worksheet.mergeCells('B17:E17');
    // worksheet.getCell('B17:E17').value = wellsfargo.installHeadingColumn1;
    worksheet.getCell('B17:E17').value = {
        'richText': [
            { 'font': { 'bold': true, 'size': 11, 'name': 'Verdana', 'family': 1 }, 'text': wellsfargo.installHeadingColumn1_D1 },
            { 'font': { 'bold': true, 'size': 11, 'color': { 'argb': 'FF0000' }, 'name': 'Verdana', 'family': 1 }, 'text': wellsfargo.installHeadingColumn1_D2 },
            { 'font': { 'bold': true, 'size': 11, 'color': { 'argb': 'FF0000' }, 'name': 'Verdana', 'family': 1 }, 'text': wellsfargo.installHeadingColumn1_D3 }
        ]
    };
    worksheet.getCell('B17:E17').font = {
        size: 11,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('B17:E17').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('B17:E17').border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };




    worksheet.mergeCells('F17');
    worksheet.getCell('F17').value = wellsfargo.installHeadingColumn2;
    worksheet.getCell('F17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('F17').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('F17').border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };


    worksheet.mergeCells('G17');
    worksheet.getCell('G17').value = wellsfargo.installHeadingColumn3;
    worksheet.getCell('G17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('G17').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('G17').border = {
        top: { style: 'thick' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };


    worksheet.mergeCells('H17');
    worksheet.getCell('H17').value = wellsfargo.installHeadingColumn4;
    worksheet.getCell('H17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('H17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('H17').border = {
        top: { style: 'thick' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('I17');
    worksheet.getCell('I17').value = wellsfargo.installHeadingColumn5;
    worksheet.getCell('I17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('I17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('I17').border = {
        top: { style: 'thick' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('J17');
    worksheet.getCell('J17').value = wellsfargo.installHeadingColumn6;
    worksheet.getCell('J17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('J17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('J17').border = {
        top: { style: 'thick' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('K17');
    worksheet.getCell('K17').value = wellsfargo.installHeadingColumn7;
    worksheet.getCell('K17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('K17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('K17').border = {
        top: { style: 'thick' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('L17');
    worksheet.getCell('L17').value = wellsfargo.installHeadingColumn8;
    worksheet.getCell('L17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('L17').border = {
        top: { style: 'thick' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('M17');
    worksheet.getCell('M17').value = wellsfargo.installHeadingColumn9;
    worksheet.getCell('M17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('M17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('M17').border = {
        top: { style: 'thick' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('O17');
    worksheet.getCell('O17').value = wellsfargo.installHeadingColumn10;
    worksheet.getCell('O17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('O17').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('O17').border = {
        // top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };


    worksheet.mergeCells('P17');
    worksheet.getCell('P17').value = wellsfargo.installHeadingColumn11;
    worksheet.getCell('P17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('P17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('P17').border = {
        // top: { style: 'thick' },
        // left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('Q17');
    worksheet.getCell('Q17').value = wellsfargo.installHeadingColumn12;
    worksheet.getCell('Q17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('Q17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('Q17').border = {
        // top: { style: 'thick' },
        // left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('R17');
    worksheet.getCell('R17').value = wellsfargo.installHeadingColumn13;
    worksheet.getCell('R17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('R17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('R17').border = {
        // top: { style: 'thick' },
        // left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S17');
    worksheet.getCell('S17').value = wellsfargo.installHeadingColumn14;
    worksheet.getCell('S17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('S17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('S17').border = {
        // top: { style: 'thick' },
        // left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T17');
    worksheet.getCell('T17').value = wellsfargo.installHeadingColumn15;
    worksheet.getCell('T17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('T17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('T17').border = {
        // top: { style: 'thick' },
        // left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('U16:U17');
    worksheet.getCell('U16:U17').value = wellsfargo.installHeadingColumn16;
    worksheet.getCell('U16:U17').font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('U16:U17').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('U16:U17').border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };


    // -----------------------------------value starting---------------------------

    for (let i = 0, j = 1.1; i < wellsfargo.installColumns.length; i++, j = j + 0.1) {
        let temp = i + 18;

        const row = worksheet.getRow(temp);
        row.height = 30;

        let cellAAlias = 'A' + temp;
        let cellBAlias = 'B' + temp;
        let cellCAlias = 'C' + temp;
        let cellDAlias = 'D' + temp;
        let cellEAlias = 'E' + temp;
        let cellFAlias = 'F' + temp;
        let cellGAlias = 'G' + temp;
        let cellHAlias = 'H' + temp;
        let cellIAlias = 'I' + temp;
        let cellJAlias = 'J' + temp;
        let cellKAlias = 'K' + temp;
        let cellLAlias = 'L' + temp;
        let cellMAlias = 'M' + temp;
        let cellNAlias = 'N' + temp;
        let cellOAlias = 'O' + temp;
        let cellPAlias = 'P' + temp;
        let cellQAlias = 'Q' + temp;
        let cellRAlias = 'R' + temp;
        let cellSAlias = 'S' + temp;
        let cellTAlias = 'T' + temp;
        let cellUAlias = 'U' + temp;


        worksheet.mergeCells(cellAAlias);
        worksheet.getCell(cellAAlias).value = j;
        worksheet.getCell(cellAAlias).font = {
            size: 11,
            name: 'Verdana',
            family: 1,
        };
        worksheet.getCell(cellAAlias).alignment = { vertical: 'middle', horizontal: 'left' };

        if (i % 2 != 1) {
            [cellBAlias, cellFAlias, cellGAlias, cellHAlias, cellIAlias, cellJAlias, cellKAlias, cellLAlias, cellMAlias,
                cellOAlias, cellPAlias, cellQAlias, cellRAlias, cellSAlias, cellTAlias, cellUAlias
            ].map(key => {
                worksheet.getCell(key).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'F2F2F2' },
                    bgColor: { argb: 'F2F2F2' }
                };
            });
        }

        worksheet.mergeCells(temp, 2, temp, 5);
        worksheet.getCell(temp, 2, temp, 5).value = wellsfargo.installColumns[i].coloumn1;
        worksheet.getCell(temp, 2, temp, 5).font = {
            size: 11,
            name: 'Calibri',
            family: 1
        };
        worksheet.getCell(temp, 2, temp, 5).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        worksheet.getCell(temp, 2, temp, 5).border = {
            top: { style: 'thin' },
            left: { style: 'thick' },
            bottom: { style: 'thin' },
            right: { style: 'thick' }
        };

        worksheet.mergeCells('F' + temp);
        worksheet.getCell('F' + temp).value = wellsfargo.installColumns[i].coloumn2;
        worksheet.getCell('F' + temp).font = {
            size: 11,
            name: 'Calibri',
            family: 1
        };
        worksheet.getCell('F' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        worksheet.getCell('F' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thick' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };


        worksheet.mergeCells('G' + temp);
        worksheet.getCell('G' + temp).value = wellsfargo.installColumns[i].coloumn3;
        worksheet.getCell('G' + temp).font = {
            size: 11,
            name: 'Calibri',
            family: 1
        };
        worksheet.getCell('G' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        worksheet.getCell('G' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };


        worksheet.mergeCells('H' + temp);
        worksheet.getCell('H' + temp).value = wellsfargo.installColumns[i].coloumn4;
        worksheet.getCell('H' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('H' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        worksheet.getCell('H' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('I' + temp);
        worksheet.getCell('I' + temp).value = wellsfargo.installColumns[i].coloumn5;
        worksheet.getCell('I' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('I' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        worksheet.getCell('I' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('J' + temp);
        worksheet.getCell('J' + temp).value = wellsfargo.installColumns[i].coloumn6;
        worksheet.getCell('J' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('J' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        worksheet.getCell('J' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('K' + temp);
        worksheet.getCell('K' + temp).value = wellsfargo.installColumns[i].coloumn7;
        worksheet.getCell('K' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('K' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        worksheet.getCell('K' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('L' + temp);
        worksheet.getCell('L' + temp).value = wellsfargo.installColumns[i].coloumn8;
        worksheet.getCell('L' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('L' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        worksheet.getCell('L' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('M' + temp);
        worksheet.getCell('M' + temp).value = wellsfargo.installColumns[i].coloumn9;
        worksheet.getCell('M' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('M' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        worksheet.getCell('M' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thick' }
        };


        worksheet.mergeCells('O' + temp);
        worksheet.getCell('O' + temp).value = wellsfargo.installColumns[i].coloumn10;
        worksheet.getCell('O' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('O' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        worksheet.getCell('O' + temp).border = {
            left: { style: 'thick' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };


        worksheet.mergeCells('P' + temp);
        worksheet.getCell('P' + temp).value = wellsfargo.installColumns[i].coloumn11;
        worksheet.getCell('P' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('P' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        worksheet.getCell('P' + temp).border = {
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('Q' + temp);
        worksheet.getCell('Q' + temp).value = wellsfargo.installColumns[i].coloumn12;
        worksheet.getCell('Q' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('Q' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        worksheet.getCell('Q' + temp).border = {
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('R' + temp);
        worksheet.getCell('R' + temp).value = "";
        worksheet.getCell('R' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('R' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        worksheet.getCell('R' + temp).border = {
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };


        worksheet.mergeCells('S' + temp);
        worksheet.getCell('S' + temp).value = wellsfargo.installColumns[i].coloumn14;
        worksheet.getCell('S' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1,
            bold: true
        };
        worksheet.getCell('S' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        worksheet.getCell('S' + temp).border = {
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('T' + temp);
        worksheet.getCell('T' + temp).value = wellsfargo.installColumns[i].coloumn15;
        worksheet.getCell('T' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('T' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        worksheet.getCell('T' + temp).border = {
            bottom: { style: 'thin' },
            right: { style: 'thick' }
        };

        worksheet.mergeCells('U' + temp);
        worksheet.getCell('U' + temp).value = wellsfargo.installColumns[i].coloumn16;
        worksheet.getCell('U' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1,
            bold: true
        };
        worksheet.getCell('U' + temp).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        worksheet.getCell('U' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thick' },
            bottom: { style: 'thin' },
            right: { style: 'thick' }
        };
    }

    // -----------------------------------value ending---------------------------

    let installColumnslength = wellsfargo.installColumns.length + 18;
    let mergeCellAlias = 'B' + installColumnslength + ':' + 'K' + installColumnslength;
    worksheet.mergeCells(mergeCellAlias);
    worksheet.mergeCells('L' + installColumnslength);
    worksheet.getCell('L' + installColumnslength).value = wellsfargo.totalProduct;
    worksheet.getCell('L' + installColumnslength).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L' + installColumnslength).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('L' + installColumnslength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('O' + installColumnslength);
    worksheet.getCell('O' + installColumnslength).value = wellsfargo.totalPeople;
    worksheet.getCell('O' + installColumnslength).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('O' + installColumnslength).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('O' + installColumnslength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('P' + installColumnslength);
    worksheet.getCell('P' + installColumnslength).value = wellsfargo.totalHoursPerPerson;
    worksheet.getCell('P' + installColumnslength).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('P' + installColumnslength).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('P' + installColumnslength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('Q' + installColumnslength);
    worksheet.getCell('Q' + installColumnslength).value = wellsfargo.totalHourlyBillRate;
    worksheet.getCell('Q' + installColumnslength).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        // bold: true
    };
    worksheet.getCell('Q' + installColumnslength).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('Q' + installColumnslength).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('Q' + installColumnslength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + installColumnslength);
    worksheet.getCell('R' + installColumnslength).value = wellsfargo.totalUnionRate;
    worksheet.getCell('R' + installColumnslength).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        // bold: true
    };
    worksheet.getCell('R' + installColumnslength).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('R' + installColumnslength).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('R' + installColumnslength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('S' + installColumnslength);
    worksheet.getCell('S' + installColumnslength).value = wellsfargo.totalLabor;
    worksheet.getCell('S' + installColumnslength).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('S' + installColumnslength).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + installColumnslength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('T' + installColumnslength);
    worksheet.getCell('T' + installColumnslength).value = wellsfargo.totalLaborTaxable;
    worksheet.getCell('T' + installColumnslength).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('T' + installColumnslength).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('T' + installColumnslength).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('T' + installColumnslength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('U' + installColumnslength);
    worksheet.getCell('U' + installColumnslength).value = wellsfargo.totalProductAndLabor;
    worksheet.getCell('U' + installColumnslength).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('U' + installColumnslength).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('U' + installColumnslength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    mergeCellAlias = 'B' + (installColumnslength + 1) + ':' + 'U' + (installColumnslength + 1);
    worksheet.mergeCells(mergeCellAlias);
    worksheet.getCell(mergeCellAlias).border = {
        top: { style: 'none' },
        left: { style: 'none' },
        bottom: { style: 'none' },
        right: { style: 'none' }
    };

    mergeCellAlias = 'B' + (installColumnslength + 2) + ':' + 'T' + (installColumnslength + 2);
    worksheet.mergeCells(mergeCellAlias);
    worksheet.getCell(mergeCellAlias).value = wellsfargo.demoHeading;
    worksheet.getCell(mergeCellAlias).font = {
        size: 16,
        name: 'Verdana',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell(mergeCellAlias).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell(mergeCellAlias).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    worksheet.getCell(mergeCellAlias).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };


    mergeCellAlias = 'B' + (installColumnslength + 3) + ':' + 'M' + (installColumnslength + 3);
    worksheet.mergeCells(mergeCellAlias);
    worksheet.getCell(mergeCellAlias).value = wellsfargo.demoSubHeadOne;
    worksheet.getCell(mergeCellAlias).font = {
        size: 16,
        name: 'Verdana',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell(mergeCellAlias).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell(mergeCellAlias).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell(mergeCellAlias).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    mergeCellAlias = 'O' + (installColumnslength + 3) + ':' + 'T' + (installColumnslength + 3);
    worksheet.mergeCells(mergeCellAlias);
    worksheet.getCell(mergeCellAlias).value = wellsfargo.demoSubHeadTwo;
    worksheet.getCell(mergeCellAlias).font = {
        size: 16,
        name: 'Verdana',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell(mergeCellAlias).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell(mergeCellAlias).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell(mergeCellAlias).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };


    // worksheet.getRow(installColumnslength+4).height = 50;
    mergeCellAlias = 'B' + (installColumnslength + 4) + ':' + 'G' + (installColumnslength + 4);
    worksheet.mergeCells(mergeCellAlias);
    worksheet.getCell(mergeCellAlias).value = wellsfargo.demoHeadingColumn1;
    worksheet.getCell(mergeCellAlias).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        color: { argb: 'FF0000' },
        bold: true
    };
    worksheet.getCell(mergeCellAlias).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell(mergeCellAlias).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    let cellAlias = 'B' + (installColumnslength + 4) + ':' + 'G' + (installColumnslength + 4);
    worksheet.mergeCells('H' + (installColumnslength + 4));
    worksheet.getCell('H' + (installColumnslength + 4)).value = "";
    worksheet.getCell('H' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('H' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('H' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('I' + (installColumnslength + 4));
    worksheet.getCell('I' + (installColumnslength + 4)).value = "";
    worksheet.getCell('I' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('I' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('I' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('J' + (installColumnslength + 4));
    worksheet.getCell('J' + (installColumnslength + 4)).value = wellsfargo.demoHeadingColumn4;
    worksheet.getCell('J' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('J' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('J' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('K' + (installColumnslength + 4));
    worksheet.getCell('K' + (installColumnslength + 4)).value = wellsfargo.demoHeadingColumn5;
    worksheet.getCell('K' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('K' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('K' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('L' + (installColumnslength + 4));
    worksheet.getCell('L' + (installColumnslength + 4)).value = wellsfargo.demoHeadingColumn6;
    worksheet.getCell('L' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('L' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('M' + (installColumnslength + 4));
    worksheet.getCell('M' + (installColumnslength + 4)).value = wellsfargo.demoHeadingColumn7;
    worksheet.getCell('M' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('M' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('M' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('O' + (installColumnslength + 4));
    worksheet.getCell('O' + (installColumnslength + 4)).value = wellsfargo.demoHeadingColumn8;
    worksheet.getCell('O' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('O' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('O' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('P' + (installColumnslength + 4));
    worksheet.getCell('P' + (installColumnslength + 4)).value = wellsfargo.demoHeadingColumn9;
    worksheet.getCell('P' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('P' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('P' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('Q' + (installColumnslength + 4));
    worksheet.getCell('Q' + (installColumnslength + 4)).value = wellsfargo.demoHeadingColumn10;
    worksheet.getCell('Q' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('Q' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('Q' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + (installColumnslength + 4));
    worksheet.getCell('R' + (installColumnslength + 4)).value = wellsfargo.demoHeadingColumn11;
    worksheet.getCell('R' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('R' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('R' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('S' + (installColumnslength + 4));
    worksheet.getCell('S' + (installColumnslength + 4)).value = wellsfargo.demoHeadingColumn12;
    worksheet.getCell('S' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('S' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('S' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('T' + (installColumnslength + 4));
    worksheet.getCell('T' + (installColumnslength + 4)).value = wellsfargo.demoHeadingColumn13;
    worksheet.getCell('T' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('T' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('T' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('U' + (installColumnslength + 2) + ':' + 'U' + (installColumnslength + 4));
    worksheet.getCell('U' + (installColumnslength + 2) + ':' + 'U' + (installColumnslength + 4)).value = wellsfargo.demoHeadingColumn14;
    worksheet.getCell('U' + (installColumnslength + 2) + ':' + 'U' + (installColumnslength + 4)).font = {
        size: 12,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('U' + (installColumnslength + 2) + ':' + 'U' + (installColumnslength + 4)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('U' + (installColumnslength + 2) + ':' + 'U' + (installColumnslength + 4)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };


    //    ----------------------------SECOND Table-----------------
    let startTableLength = installColumnslength + 4 + 1;
    let demoTableLength = startTableLength + wellsfargo.demoColumns.length;
    for (let i = startTableLength, j = 20.1, k = 0; i < demoTableLength; i++, j = j + 0.1, k++) {
        let temp = i;

        const row = worksheet.getRow(temp);
        // row.height = 30;

        let cellAAlias = 'A' + temp;
        let cellBAlias = 'B' + temp;
        let cellCAlias = 'C' + temp;
        let cellDAlias = 'D' + temp;
        let cellEAlias = 'E' + temp;
        let cellFAlias = 'F' + temp;
        let cellGAlias = 'G' + temp;
        let cellHAlias = 'H' + temp;
        let cellIAlias = 'I' + temp;
        let cellJAlias = 'J' + temp;
        let cellKAlias = 'K' + temp;
        let cellLAlias = 'L' + temp;
        let cellMAlias = 'M' + temp;
        let cellNAlias = 'N' + temp;
        let cellOAlias = 'O' + temp;
        let cellPAlias = 'P' + temp;
        let cellQAlias = 'Q' + temp;
        let cellRAlias = 'R' + temp;
        let cellSAlias = 'S' + temp;
        let cellTAlias = 'T' + temp;
        let cellUAlias = 'U' + temp;

        worksheet.mergeCells(cellAAlias);
        worksheet.getCell(cellAAlias).value = j;
        worksheet.getCell(cellAAlias).font = {
            size: 11,
            name: 'Verdana',
            family: 1,
        };
        worksheet.getCell(cellAAlias).alignment = { vertical: 'middle', horizontal: 'left' };

        if (i % 2 != 1) {
            [cellBAlias, cellFAlias, cellGAlias, cellHAlias, cellIAlias, cellJAlias, cellKAlias, cellLAlias, cellMAlias,
                cellOAlias, cellPAlias, cellQAlias, cellRAlias, cellSAlias, cellTAlias, cellUAlias
            ].map(key => {
                worksheet.getCell(key).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'F2F2F2' },
                    bgColor: { argb: 'F2F2F2' }
                };
            });
        }

        worksheet.mergeCells(temp, 2, temp, 7);
        worksheet.getCell(temp, 2, temp, 7).value = wellsfargo.demoColumns[k].coloumn1;
        worksheet.getCell(temp, 2, temp, 7).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell(temp, 2, temp, 7).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        worksheet.getCell(temp, 2, temp, 7).border = {
            top: { style: 'thin' },
            left: { style: 'thick' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('H' + temp);
        worksheet.getCell('H' + temp).value = wellsfargo.demoColumns[k].coloumn2;
        worksheet.getCell('H' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };

        worksheet.getCell('H' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        worksheet.getCell('H' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('I' + temp);
        worksheet.getCell('I' + temp).value = wellsfargo.demoColumns[k].coloumn3;
        worksheet.getCell('I' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('I' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        worksheet.getCell('I' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('J' + temp);
        worksheet.getCell('J' + temp).value = wellsfargo.demoColumns[k].coloumn4;
        worksheet.getCell('J' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('J' + temp).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        worksheet.getCell('J' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('K' + temp);
        worksheet.getCell('K' + temp).value = wellsfargo.demoColumns[k].coloumn5;
        worksheet.getCell('K' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('K' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        worksheet.getCell('K' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('L' + temp);
        worksheet.getCell('L' + temp).value = wellsfargo.demoColumns[k].coloumn6;
        worksheet.getCell('L' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('L' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        worksheet.getCell('L' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('M' + temp);
        worksheet.getCell('M' + temp).value = wellsfargo.demoColumns[k].coloumn7;
        worksheet.getCell('M' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('M' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        worksheet.getCell('M' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thick' }
        };

        worksheet.mergeCells('O' + temp);
        worksheet.getCell('O' + temp).value = wellsfargo.demoColumns[k].coloumn8;
        worksheet.getCell('O' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('O' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        worksheet.getCell('O' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thick' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('P' + temp);
        worksheet.getCell('P' + temp).value = wellsfargo.demoColumns[k].coloumn9;
        worksheet.getCell('P' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('P' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        worksheet.getCell('P' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };


        worksheet.mergeCells('Q' + temp);
        worksheet.getCell('Q' + temp).value = wellsfargo.demoColumns[k].coloumn10;
        worksheet.getCell('Q' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('Q' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        worksheet.getCell('Q' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('R' + temp);
        worksheet.getCell('R' + temp).value = wellsfargo.demoColumns[k].coloumn11;
        worksheet.getCell('R' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('R' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        worksheet.getCell('R' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('S' + temp);
        worksheet.getCell('S' + temp).value = wellsfargo.demoColumns[k].coloumn12;
        worksheet.getCell('S' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('S' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        worksheet.getCell('S' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('T' + temp);
        worksheet.getCell('T' + temp).value = wellsfargo.demoColumns[k].coloumn13;
        worksheet.getCell('T' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('T' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        worksheet.getCell('T' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        worksheet.mergeCells('U' + temp);
        worksheet.getCell('U' + temp).value = wellsfargo.demoColumns[k].coloumn14;
        worksheet.getCell('U' + temp).font = {
            size: 13,
            name: 'Verdana',
            family: 1
        };
        worksheet.getCell('U' + temp).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        worksheet.getCell('U' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thick' }
        };

    }


    //    ----------------------------SECOND ROW-----------------

    mergeCellAlias = 'B' + demoTableLength + ':' + 'K' + demoTableLength;
    worksheet.mergeCells(mergeCellAlias);

    worksheet.mergeCells('L' + demoTableLength);
    worksheet.getCell('L' + demoTableLength).value = wellsfargo.totalDemo;
    worksheet.getCell('L' + demoTableLength).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L' + demoTableLength).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('L' + demoTableLength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('O' + demoTableLength);
    worksheet.getCell('O' + demoTableLength).value = wellsfargo.totalDemoPeople;
    worksheet.getCell('O' + demoTableLength).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('O' + demoTableLength).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('O' + demoTableLength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('P' + demoTableLength);
    worksheet.getCell('P' + demoTableLength).value = wellsfargo.totalDemoHoursPerPerson;
    worksheet.getCell('P' + demoTableLength).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('P' + demoTableLength).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('P' + demoTableLength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('Q' + demoTableLength);
    worksheet.getCell('Q' + demoTableLength).value = wellsfargo.totalDemoHourlyBillRate;
    worksheet.getCell('Q' + demoTableLength).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('Q' + demoTableLength).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('Q' + demoTableLength).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('Q' + demoTableLength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + demoTableLength);
    worksheet.getCell('R' + demoTableLength).value = wellsfargo.totalDemoUnionRate;
    worksheet.getCell('R' + demoTableLength).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('R' + demoTableLength).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('R' + demoTableLength).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + demoTableLength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('S' + demoTableLength);
    worksheet.getCell('S' + demoTableLength).value = wellsfargo.totalDemoLabor;
    worksheet.getCell('S' + demoTableLength).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('S' + demoTableLength).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + demoTableLength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('T' + demoTableLength);
    worksheet.getCell('T' + demoTableLength).value = wellsfargo.totalDemoLaborTaxable;
    worksheet.getCell('T' + demoTableLength).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('T' + demoTableLength).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('T' + demoTableLength).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + demoTableLength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('U' + demoTableLength);
    worksheet.getCell('U' + demoTableLength).value = wellsfargo.totalDemoProductAndLabor;
    worksheet.getCell('U' + demoTableLength).font = {
        size: 13,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('U' + demoTableLength).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('U' + demoTableLength).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };



    worksheet.mergeCells('B' + (demoTableLength + 1) + ':' + 'U' + (demoTableLength + 1));
    worksheet.getCell('B' + (demoTableLength + 1) + ':' + 'U' + (demoTableLength + 1)).border = {
        top: { style: 'none' },
        left: { style: 'none' },
        bottom: { style: 'none' },
        right: { style: 'none' }
    };


    worksheet.mergeCells('B' + (demoTableLength + 2) + ':' + 'M' + (demoTableLength + 2));
    worksheet.getCell('B' + (demoTableLength + 2) + ':' + 'M' + (demoTableLength + 2)).value = wellsfargo.clarificationsHeading;
    worksheet.getCell('B' + (demoTableLength + 2) + ':' + 'M' + (demoTableLength + 2)).font = {
        size: 16,
        name: 'Verdana',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell('B' + (demoTableLength + 2) + ':' + 'M' + (demoTableLength + 2)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '808080' },
        bgColor: { argb: '808080' }
    };
    worksheet.getCell('B' + (demoTableLength + 2) + ':' + 'M' + (demoTableLength + 2)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('B' + (demoTableLength + 2) + ':' + 'M' + (demoTableLength + 2)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.getRow(demoTableLength + 3).height = 120;
    worksheet.mergeCells('A' + (demoTableLength + 3));
    worksheet.getCell('A' + (demoTableLength + 3)).value = "30.1";
    worksheet.getCell('A' + (demoTableLength + 3)).alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('A' + (demoTableLength + 3)).font = {
        size: 11,
        name: 'Verdana',
        family: 1
    };


    worksheet.mergeCells('B' + (demoTableLength + 3) + ':' + 'M' + (demoTableLength + 3));
    worksheet.getCell('B' + (demoTableLength + 3) + ':' + 'M' + (demoTableLength + 3)).value = wellsfargo.clarificationDescription;
    worksheet.getCell('B' + (demoTableLength + 3) + ':' + 'M' + (demoTableLength + 3)).font = {
        size: 16,
        name: 'Verdana',
        family: 1
    };
    worksheet.getCell('B' + (demoTableLength + 3) + ':' + 'M' + (demoTableLength + 3)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('B' + (demoTableLength + 3) + ':' + 'M' + (demoTableLength + 3)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };




    worksheet.mergeCells('S' + (demoTableLength + 4) + ':' + 'U' + (demoTableLength + 4));
    worksheet.getCell('S' + (demoTableLength + 4) + ':' + 'U' + (demoTableLength + 4)).value = wellsfargo.taxHeading;
    worksheet.getCell('S' + (demoTableLength + 4) + ':' + 'U' + (demoTableLength + 4)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('S' + (demoTableLength + 4) + ':' + 'U' + (demoTableLength + 4)).alignment = { vertical: 'bottom', horizontal: 'right', wrapText: true };



    worksheet.mergeCells('B' + (demoTableLength + 5) + ':' + 'K' + (demoTableLength + 5));

    worksheet.mergeCells('L' + (demoTableLength + 5) + ':' + 'M' + (demoTableLength + 5));
    worksheet.getCell('L' + (demoTableLength + 5) + ':' + 'M' + (demoTableLength + 5)).value = wellsfargo.taxHeadingColoumn1;
    worksheet.getCell('L' + (demoTableLength + 5) + ':' + 'M' + (demoTableLength + 5)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L' + (demoTableLength + 5) + ':' + 'M' + (demoTableLength + 5)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('L' + (demoTableLength + 5) + ':' + 'M' + (demoTableLength + 5)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('O' + (demoTableLength + 5) + ':' + 'Q' + (demoTableLength + 5));

    worksheet.mergeCells('R' + (demoTableLength + 5));
    worksheet.getCell('R' + (demoTableLength + 5)).value = wellsfargo.taxHeadingColoumn2;
    worksheet.getCell('R' + (demoTableLength + 5)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('R' + (demoTableLength + 5)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('R' + (demoTableLength + 5)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('S' + (demoTableLength + 5));
    worksheet.getCell('S' + (demoTableLength + 5)).value = wellsfargo.taxHeadingColoumn3;
    worksheet.getCell('S' + (demoTableLength + 5)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('S' + (demoTableLength + 5)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('S' + (demoTableLength + 5)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('T' + (demoTableLength + 5) + ':' + 'U' + (demoTableLength + 5));
    worksheet.getCell('T' + (demoTableLength + 5) + ':' + 'U' + (demoTableLength + 5)).value = wellsfargo.totalTaxHeading;
    worksheet.getCell('T' + (demoTableLength + 5) + ':' + 'U' + (demoTableLength + 5)).font = {
        size: 16,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('T' + (demoTableLength + 5) + ':' + 'U' + (demoTableLength + 5)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('T' + (demoTableLength + 5) + ':' + 'U' + (demoTableLength + 5)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('A' + (demoTableLength + 6));
    worksheet.getCell('A' + (demoTableLength + 6)).value = "40";
    worksheet.getCell('A' + (demoTableLength + 6)).alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('A' + (demoTableLength + 6)).font = {
        size: 11,
        name: 'Verdana',
        family: 1,
    };
    worksheet.mergeCells('L' + (demoTableLength + 6) + ':' + 'M' + (demoTableLength + 6));
    worksheet.getCell('L' + (demoTableLength + 6) + ':' + 'M' + (demoTableLength + 6)).value = wellsfargo.taxRate;
    worksheet.getCell('L' + (demoTableLength + 6) + ':' + 'M' + (demoTableLength + 6)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('L' + (demoTableLength + 6) + ':' + 'M' + (demoTableLength + 6)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('L' + (demoTableLength + 6) + ':' + 'M' + (demoTableLength + 6)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    worksheet.getCell('L' + (demoTableLength + 6) + ':' + 'M' + (demoTableLength + 6)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('O' + (demoTableLength + 6) + ':' + 'Q' + (demoTableLength + 6));
    worksheet.getCell('O' + (demoTableLength + 6) + ':' + 'Q' + (demoTableLength + 6)).value = wellsfargo.product;
    worksheet.getCell('O' + (demoTableLength + 6) + ':' + 'Q' + (demoTableLength + 6)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('O' + (demoTableLength + 6) + ':' + 'Q' + (demoTableLength + 6)).alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
    worksheet.getCell('O' + (demoTableLength + 6) + ':' + 'Q' + (demoTableLength + 6)).border = {
        top: { style: 'thick' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + (demoTableLength + 6));
    worksheet.getCell('R' + (demoTableLength + 6)).value = wellsfargo.productPreTax;
    worksheet.getCell('R' + (demoTableLength + 6)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('R' + (demoTableLength + 6)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + (demoTableLength + 6)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S' + (demoTableLength + 6));
    worksheet.getCell('S' + (demoTableLength + 6)).value = wellsfargo.productTax;
    worksheet.getCell('S' + (demoTableLength + 6)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('S' + (demoTableLength + 6)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('S' + (demoTableLength + 6)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + (demoTableLength + 6)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + (demoTableLength + 6) + ':' + 'U' + (demoTableLength + 6));
    worksheet.getCell('T' + (demoTableLength + 6) + ':' + 'U' + (demoTableLength + 6)).value = wellsfargo.productTotal;
    worksheet.getCell('T' + (demoTableLength + 6) + ':' + 'U' + (demoTableLength + 6)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('T' + (demoTableLength + 6) + ':' + 'U' + (demoTableLength + 6)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + (demoTableLength + 6) + ':' + 'U' + (demoTableLength + 6)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('A' + (demoTableLength + 7));
    worksheet.getCell('A' + (demoTableLength + 7)).value = "40";
    worksheet.getCell('A' + (demoTableLength + 7)).alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('A' + (demoTableLength + 7)).font = {
        size: 11,
        name: 'Verdana',
        family: 1
    };

    worksheet.mergeCells('O' + (demoTableLength + 7) + ':' + 'Q' + (demoTableLength + 7));
    worksheet.getCell('O' + (demoTableLength + 7) + ':' + 'Q' + (demoTableLength + 7)).value = wellsfargo.labor;
    worksheet.getCell('O' + (demoTableLength + 7) + ':' + 'Q' + (demoTableLength + 7)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('O' + (demoTableLength + 7) + ':' + 'Q' + (demoTableLength + 7)).alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
    worksheet.getCell('O' + (demoTableLength + 7) + ':' + 'Q' + (demoTableLength + 7)).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + (demoTableLength + 7));
    worksheet.getCell('R' + (demoTableLength + 7)).value = wellsfargo.laborPreTax;
    worksheet.getCell('R' + (demoTableLength + 7)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('R' + (demoTableLength + 7)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + (demoTableLength + 7)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S' + (demoTableLength + 7));
    worksheet.getCell('S' + (demoTableLength + 7)).value = wellsfargo.laborTax;
    worksheet.getCell('S' + (demoTableLength + 7)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('S' + (demoTableLength + 7)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('S' + (demoTableLength + 7)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + (demoTableLength + 7)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + (demoTableLength + 7) + ':' + 'U' + (demoTableLength + 7));
    worksheet.getCell('T' + (demoTableLength + 7) + ':' + 'U' + (demoTableLength + 7)).value = wellsfargo.laborTotal;
    worksheet.getCell('T' + (demoTableLength + 7) + ':' + 'U' + (demoTableLength + 7)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('T' + (demoTableLength + 7) + ':' + 'U' + (demoTableLength + 7)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + (demoTableLength + 7) + ':' + 'U' + (demoTableLength + 7)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };



    worksheet.mergeCells('A' + (demoTableLength + 8));
    worksheet.getCell('A' + (demoTableLength + 8)).value = "40";
    worksheet.getCell('A' + (demoTableLength + 8)).alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('A' + (demoTableLength + 8)).font = {
        size: 11,
        name: 'Verdana',
        family: 1,
    };

    worksheet.mergeCells('O' + (demoTableLength + 8) + ':' + 'Q' + (demoTableLength + 8));
    worksheet.getCell('O' + (demoTableLength + 8) + ':' + 'Q' + (demoTableLength + 8)).value = wellsfargo.subTotal1;
    worksheet.getCell('O' + (demoTableLength + 8) + ':' + 'Q' + (demoTableLength + 8)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('O' + (demoTableLength + 8) + ':' + 'Q' + (demoTableLength + 8)).alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
    worksheet.getCell('O' + (demoTableLength + 8) + ':' + 'Q' + (demoTableLength + 8)).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + (demoTableLength + 8));
    worksheet.getCell('R' + (demoTableLength + 8)).value = wellsfargo.subTotal1PreTax;
    worksheet.getCell('R' + (demoTableLength + 8)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('R' + (demoTableLength + 8)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + (demoTableLength + 8)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S' + (demoTableLength + 8));
    worksheet.getCell('S' + (demoTableLength + 8)).value = wellsfargo.subTotal1Tax;
    worksheet.getCell('S' + (demoTableLength + 8)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('S' + (demoTableLength + 8)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + (demoTableLength + 8)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + (demoTableLength + 8) + ':' + 'U' + (demoTableLength + 8));
    worksheet.getCell('T' + (demoTableLength + 8) + ':' + 'U' + (demoTableLength + 8)).value = wellsfargo.subTotal1Total;
    worksheet.getCell('T' + (demoTableLength + 8) + ':' + 'U' + (demoTableLength + 8)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('T' + (demoTableLength + 8) + ':' + 'U' + (demoTableLength + 8)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + (demoTableLength + 8) + ':' + 'U' + (demoTableLength + 8)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('A' + (demoTableLength + 9));
    worksheet.getCell('A' + (demoTableLength + 9)).value = "40";
    worksheet.getCell('A' + (demoTableLength + 9)).alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('A' + (demoTableLength + 9)).font = {
        size: 11,
        name: 'Verdana',
        family: 1
    };

    worksheet.mergeCells('O' + (demoTableLength + 9) + ':' + 'Q' + (demoTableLength + 9));
    worksheet.getCell('O' + (demoTableLength + 9) + ':' + 'Q' + (demoTableLength + 9)).value = wellsfargo.demoProduct;
    worksheet.getCell('O' + (demoTableLength + 9) + ':' + 'Q' + (demoTableLength + 9)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('O' + (demoTableLength + 9) + ':' + 'Q' + (demoTableLength + 9)).alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
    worksheet.getCell('O' + (demoTableLength + 9) + ':' + 'Q' + (demoTableLength + 9)).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + (demoTableLength + 9));
    worksheet.getCell('R' + (demoTableLength + 9)).value = wellsfargo.demoProductPreTax;
    worksheet.getCell('R' + (demoTableLength + 9)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('R' + (demoTableLength + 9)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + (demoTableLength + 9)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S' + (demoTableLength + 9));
    worksheet.getCell('S' + (demoTableLength + 9)).value = wellsfargo.demoProductTax;
    worksheet.getCell('S' + (demoTableLength + 9)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('S' + (demoTableLength + 9)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('S' + (demoTableLength + 9)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + (demoTableLength + 9)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + (demoTableLength + 9) + ':' + 'U' + (demoTableLength + 9));
    worksheet.getCell('T' + (demoTableLength + 9) + ':' + 'U' + (demoTableLength + 9)).value = wellsfargo.demoProductTotal;
    worksheet.getCell('T' + (demoTableLength + 9) + ':' + 'U' + (demoTableLength + 9)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('T' + (demoTableLength + 9) + ':' + 'U' + (demoTableLength + 9)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + (demoTableLength + 9) + ':' + 'U' + (demoTableLength + 9)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('A' + (demoTableLength + 10));
    worksheet.getCell('A' + (demoTableLength + 10)).value = "41";
    worksheet.getCell('A' + (demoTableLength + 10)).alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('A' + (demoTableLength + 10)).font = {
        size: 11,
        name: 'Verdana',
        family: 1
    };

    worksheet.mergeCells('O' + (demoTableLength + 10) + ':' + 'Q' + (demoTableLength + 10));
    worksheet.getCell('O' + (demoTableLength + 10) + ':' + 'Q' + (demoTableLength + 10)).value = wellsfargo.demoLabor;
    worksheet.getCell('O' + (demoTableLength + 10) + ':' + 'Q' + (demoTableLength + 10)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('O' + (demoTableLength + 10) + ':' + 'Q' + (demoTableLength + 10)).alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
    worksheet.getCell('O' + (demoTableLength + 10) + ':' + 'Q' + (demoTableLength + 10)).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + (demoTableLength + 10));
    worksheet.getCell('R' + (demoTableLength + 10)).value = wellsfargo.demoLaborPreTax;
    worksheet.getCell('R' + (demoTableLength + 10)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('R' + (demoTableLength + 10)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + (demoTableLength + 10)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S' + (demoTableLength + 10));
    worksheet.getCell('S' + (demoTableLength + 10)).value = wellsfargo.demoLaborTax;
    worksheet.getCell('S' + (demoTableLength + 10)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('S' + (demoTableLength + 10)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('S' + (demoTableLength + 10)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + (demoTableLength + 10)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + (demoTableLength + 10) + ':' + 'U' + (demoTableLength + 10));
    worksheet.getCell('T' + (demoTableLength + 10) + ':' + 'U' + (demoTableLength + 10)).value = wellsfargo.demoLaborTotal;
    worksheet.getCell('T' + (demoTableLength + 10) + ':' + 'U' + (demoTableLength + 10)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('T' + (demoTableLength + 10) + ':' + 'U' + (demoTableLength + 10)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + (demoTableLength + 10) + ':' + 'U' + (demoTableLength + 10)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('A' + (demoTableLength + 11));
    worksheet.getCell('A' + (demoTableLength + 11)).value = "41";
    worksheet.getCell('A' + (demoTableLength + 11)).alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('A' + (demoTableLength + 11)).font = {
        size: 11,
        name: 'Verdana',
        family: 1
    };
    // // ------------------------------------------.....to be continued----------------------------
    worksheet.mergeCells('O' + (demoTableLength + 11) + ':' + 'Q' + (demoTableLength + 11));
    worksheet.getCell('O' + (demoTableLength + 11) + ':' + 'Q' + (demoTableLength + 11)).value = wellsfargo.subTotal2;
    worksheet.getCell('O' + (demoTableLength + 11) + ':' + 'Q' + (demoTableLength + 11)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('O' + (demoTableLength + 11) + ':' + 'Q' + (demoTableLength + 11)).alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
    worksheet.getCell('O' + (demoTableLength + 11) + ':' + 'Q' + (demoTableLength + 11)).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + (demoTableLength + 11));
    worksheet.getCell('R' + (demoTableLength + 11)).value = wellsfargo.subTotal2PreTax;
    worksheet.getCell('R' + (demoTableLength + 11)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('R' + (demoTableLength + 11)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + (demoTableLength + 11)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S' + (demoTableLength + 11));
    worksheet.getCell('S' + (demoTableLength + 11)).value = wellsfargo.subTotal2Tax;
    worksheet.getCell('S' + (demoTableLength + 11)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('S' + (demoTableLength + 11)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + (demoTableLength + 11)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + (demoTableLength + 11) + ':' + 'U' + (demoTableLength + 11));
    worksheet.getCell('T' + (demoTableLength + 11) + ':' + 'U' + (demoTableLength + 11)).value = wellsfargo.subTotal2Total;
    worksheet.getCell('T' + (demoTableLength + 11) + ':' + 'U' + (demoTableLength + 11)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('T' + (demoTableLength + 11) + ':' + 'U' + (demoTableLength + 11)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + (demoTableLength + 11) + ':' + 'U' + (demoTableLength + 11)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('O' + (demoTableLength + 12) + ':' + 'Q' + (demoTableLength + 12));
    worksheet.getCell('O' + (demoTableLength + 12) + ':' + 'Q' + (demoTableLength + 12)).value = wellsfargo.freight;
    worksheet.getCell('O' + (demoTableLength + 12) + ':' + 'Q' + (demoTableLength + 12)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('O' + (demoTableLength + 12) + ':' + 'Q' + (demoTableLength + 12)).alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
    worksheet.getCell('O' + (demoTableLength + 12) + ':' + 'Q' + (demoTableLength + 12)).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + (demoTableLength + 12));
    worksheet.getCell('R' + (demoTableLength + 12)).value = wellsfargo.freightPreTax;
    worksheet.getCell('R' + (demoTableLength + 12)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('R' + (demoTableLength + 12)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('R' + (demoTableLength + 12)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + (demoTableLength + 12)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S' + (demoTableLength + 12));
    worksheet.getCell('S' + (demoTableLength + 12)).value = wellsfargo.freightTax;
    worksheet.getCell('S' + (demoTableLength + 12)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('S' + (demoTableLength + 12)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('S' + (demoTableLength + 12)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + (demoTableLength + 12)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + (demoTableLength + 12) + ':' + 'U' + (demoTableLength + 12));
    worksheet.getCell('T' + (demoTableLength + 12) + ':' + 'U' + (demoTableLength + 12)).value = wellsfargo.freightTotal;
    worksheet.getCell('T' + (demoTableLength + 12) + ':' + 'U' + (demoTableLength + 12)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('T' + (demoTableLength + 12) + ':' + 'U' + (demoTableLength + 12)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + (demoTableLength + 12) + ':' + 'U' + (demoTableLength + 12)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('O' + (demoTableLength + 13) + ':' + 'Q' + (demoTableLength + 13));
    worksheet.getCell('O' + (demoTableLength + 13) + ':' + 'Q' + (demoTableLength + 13)).value = wellsfargo.shipping;
    worksheet.getCell('O' + (demoTableLength + 13) + ':' + 'Q' + (demoTableLength + 13)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('O' + (demoTableLength + 13) + ':' + 'Q' + (demoTableLength + 13)).alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
    worksheet.getCell('O' + (demoTableLength + 13) + ':' + 'Q' + (demoTableLength + 13)).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + (demoTableLength + 13));
    worksheet.getCell('R' + (demoTableLength + 13)).value = wellsfargo.shippingPreTax;
    worksheet.getCell('R' + (demoTableLength + 13)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('R' + (demoTableLength + 13)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('R' + (demoTableLength + 13)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + (demoTableLength + 13)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S' + (demoTableLength + 13));
    worksheet.getCell('S' + (demoTableLength + 13)).value = wellsfargo.shippingTax;
    worksheet.getCell('S' + (demoTableLength + 13)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('S' + (demoTableLength + 13)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('S' + (demoTableLength + 13)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + (demoTableLength + 13)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + (demoTableLength + 13) + ':' + 'U' + (demoTableLength + 13));
    worksheet.getCell('T' + (demoTableLength + 13) + ':' + 'U' + (demoTableLength + 13)).value = wellsfargo.shippingTotal;
    worksheet.getCell('T' + (demoTableLength + 13) + ':' + 'U' + (demoTableLength + 13)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('T' + (demoTableLength + 13) + ':' + 'U' + (demoTableLength + 13)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + (demoTableLength + 13) + ':' + 'U' + (demoTableLength + 13)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('O' + (demoTableLength + 14) + ':' + 'Q' + (demoTableLength + 14));
    worksheet.getCell('O' + (demoTableLength + 14) + ':' + 'Q' + (demoTableLength + 14)).value = wellsfargo.profit;
    worksheet.getCell('O' + (demoTableLength + 14) + ':' + 'Q' + (demoTableLength + 14)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('O' + (demoTableLength + 14) + ':' + 'Q' + (demoTableLength + 14)).alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
    worksheet.getCell('O' + (demoTableLength + 14) + ':' + 'Q' + (demoTableLength + 14)).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + (demoTableLength + 14));
    worksheet.getCell('R' + (demoTableLength + 14)).value = wellsfargo.profitPreTax;
    worksheet.getCell('R' + (demoTableLength + 14)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('R' + (demoTableLength + 14)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('R' + (demoTableLength + 14)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + (demoTableLength + 14)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S' + (demoTableLength + 14));
    worksheet.getCell('S' + (demoTableLength + 14)).value = wellsfargo.profitTax;
    worksheet.getCell('S' + (demoTableLength + 14)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('S' + (demoTableLength + 14)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('S' + (demoTableLength + 14)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + (demoTableLength + 14)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + (demoTableLength + 14) + ':' + 'U' + (demoTableLength + 14));
    worksheet.getCell('T' + (demoTableLength + 14) + ':' + 'U' + (demoTableLength + 14)).value = wellsfargo.profitTotal;
    worksheet.getCell('T' + (demoTableLength + 14) + ':' + 'U' + (demoTableLength + 14)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('T' + (demoTableLength + 14) + ':' + 'U' + (demoTableLength + 14)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + (demoTableLength + 14) + ':' + 'U' + (demoTableLength + 14)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };


    worksheet.mergeCells('O' + (demoTableLength + 15) + ':' + 'Q' + (demoTableLength + 15));
    worksheet.getCell('O' + (demoTableLength + 15) + ':' + 'Q' + (demoTableLength + 15)).value = wellsfargo.insurance;
    worksheet.getCell('O' + (demoTableLength + 15) + ':' + 'Q' + (demoTableLength + 15)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('O' + (demoTableLength + 15) + ':' + 'Q' + (demoTableLength + 15)).alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
    worksheet.getCell('O' + (demoTableLength + 15) + ':' + 'Q' + (demoTableLength + 15)).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + (demoTableLength + 15));
    worksheet.getCell('R' + (demoTableLength + 15)).value = wellsfargo.insurancePreTax;
    worksheet.getCell('R' + (demoTableLength + 15)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('R' + (demoTableLength + 15)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('R' + (demoTableLength + 15)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + (demoTableLength + 15)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S' + (demoTableLength + 15));
    worksheet.getCell('S' + (demoTableLength + 15)).value = wellsfargo.insuranceTax;
    worksheet.getCell('S' + (demoTableLength + 15)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2DCDB' },
        bgColor: { argb: 'F2DCDB' }
    };
    worksheet.getCell('S' + (demoTableLength + 15)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('S' + (demoTableLength + 15)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + (demoTableLength + 15)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + (demoTableLength + 15) + ':' + 'U' + (demoTableLength + 15));
    worksheet.getCell('T' + (demoTableLength + 15) + ':' + 'U' + (demoTableLength + 15)).value = wellsfargo.insuranceTotal;
    worksheet.getCell('T' + (demoTableLength + 15) + ':' + 'U' + (demoTableLength + 15)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
    };
    worksheet.getCell('T' + (demoTableLength + 15) + ':' + 'U' + (demoTableLength + 15)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + (demoTableLength + 15) + ':' + 'U' + (demoTableLength + 15)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thick' }
    };



    worksheet.mergeCells('A' + (demoTableLength + 16));
    worksheet.getCell('A' + (demoTableLength + 16)).value = "41";
    worksheet.getCell('A' + (demoTableLength + 16)).alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('A' + (demoTableLength + 16)).font = {
        size: 11,
        name: 'Verdana',
        family: 1,
    };

    worksheet.mergeCells('O' + (demoTableLength + 16) + ':' + 'Q' + (demoTableLength + 16));
    worksheet.getCell('O' + (demoTableLength + 16) + ':' + 'Q' + (demoTableLength + 16)).value = wellsfargo.allTotal;
    worksheet.getCell('O' + (demoTableLength + 16) + ':' + 'Q' + (demoTableLength + 16)).font = {
        size: 14,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('O' + (demoTableLength + 16) + ':' + 'Q' + (demoTableLength + 16)).alignment = { vertical: 'top', horizontal: 'right', wrapText: true };
    worksheet.getCell('O' + (demoTableLength + 16) + ':' + 'Q' + (demoTableLength + 16)).border = {
        top: { style: 'thin' },
        left: { style: 'thick' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('R' + (demoTableLength + 16));
    worksheet.getCell('R' + (demoTableLength + 16)).value = wellsfargo.allTotalTaxPreTax;
    worksheet.getCell('R' + (demoTableLength + 16)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('R' + (demoTableLength + 16)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('R' + (demoTableLength + 16)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('S' + (demoTableLength + 16));
    worksheet.getCell('S' + (demoTableLength + 16)).value = wellsfargo.allTotalTax;
    worksheet.getCell('S' + (demoTableLength + 16)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('S' + (demoTableLength + 16)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('S' + (demoTableLength + 16)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('T' + (demoTableLength + 16) + ':' + 'U' + (demoTableLength + 16));
    worksheet.getCell('T' + (demoTableLength + 16) + ':' + 'U' + (demoTableLength + 16)).value = wellsfargo.allMaterialTotal;
    worksheet.getCell('T' + (demoTableLength + 16) + ':' + 'U' + (demoTableLength + 16)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('T' + (demoTableLength + 16) + ':' + 'U' + (demoTableLength + 16)).alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('T' + (demoTableLength + 16) + ':' + 'U' + (demoTableLength + 16)).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thick' },
        right: { style: 'thick' }
    };

    worksheet.mergeCells('M' + (demoTableLength + 11) + ':' + 'N' + (demoTableLength + 15));
    worksheet.getCell('M' + (demoTableLength + 11) + ':' + 'N' + (demoTableLength + 15)).value = wellsfargo.rotationText;
    worksheet.getCell('M' + (demoTableLength + 11) + ':' + 'N' + (demoTableLength + 15)).font = {
        size: 13,
        name: 'Verdana',
        family: 1,
        bold: true
    };
    worksheet.getCell('M' + (demoTableLength + 11) + ':' + 'N' + (demoTableLength + 15)).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
        bgColor: { argb: 'F2F2F2' }
    };
    worksheet.getCell('M' + (demoTableLength + 11) + ':' + 'N' + (demoTableLength + 15)).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true, textRotation: 90 };


    res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
        "Content-Disposition",
        "attachment; filename=" + "wellfargo" + ".xlsx"
    );
    return workbook.xlsx.write(res).then(function() {
        res['status'](200).end();
    });

});

app.listen(PORT, function() {
    console.log('App listening on port ' + PORT);
});