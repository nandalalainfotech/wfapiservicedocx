const cors = require('cors');
const path = require('path');
const excel = require('exceljs');
const saveAs = require('file-saver');
const fs = require('fs');
var pdf = require('dynamic-html-pdf');
var html = fs.readFileSync('./template/document.html', 'utf8');
const docx = require('docx');
const {
    AlignmentType,
    convertInchesToTwip,
    TabStopPosition,
    LevelFormat,
    HeadingLevel,
    UnderlineType,
    ImageRun,
    VerticalPositionAlign,
    VerticalAlign,
    TextDirection,
    VerticalPositionRelativeFrom,
    PageNumber,
    NumberFormat,
    InternalHyperlink,
    Bookmark,
    PageReference,
    Media,
    HorizontalPositionAlign,
    PageBreak,
    FootnoteReferenceRun,
    WidthType,
    ShadingType,
    HorizontalPositionRelativeFrom,
    BorderStyle,
    PageBorderDisplay,
    PageBorderOffsetFrom,
    PageBorderZOrder,
    Table,
    TableRow,
    Footer,
    LineNumberRestartFormat,
    SequentialIdentifier,
    TextWrappingSide,
    TextWrappingType,
    RelativeHorizontalPosition,
    TableOfContents,
    StyleLevel,
    ExternalHyperlink,
    RelativeVerticalPosition,
    OverlapType,
    TableAnchorType,
    TableCell,
    TableLayoutType,
    PageNumberSeparator,
    Header,
    Document,
    Packer,
    Paragraph,
    LineRuleType,
    convertMillimetersToTwip,
    PageOrientation,
    TextRun,
    Style
} = require('docx');

const bodyParser = require("body-parser");


var express = require('express');
var app = express();

app.use(cors());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static('public'));


var PORT = process.env.PORT || 80;


const LOREM_IPSUM =
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam mi velit, convallis convallis scelerisque nec, faucibus nec leo. Phasellus at posuere mauris, tempus dignissim velit. Integer et tortor dolor. Duis auctor efficitur mattis. Vivamus ut metus accumsan tellus auctor sollicitudin venenatis et nibh. Cras quis massa ac metus fringilla venenatis. Proin rutrum mauris purus, ut suscipit magna consectetur id. Integer consectetur sollicitudin ante, vitae faucibus neque efficitur in. Praesent ultricies nibh lectus. Mauris pharetra id odio eget iaculis. Duis dictum, risus id pellentesque rutrum, lorem quam malesuada massa, quis ullamcorper turpis urna a diam. Cras vulputate metus vel massa porta ullamcorper. Etiam porta condimentum nulla nec tristique. Sed nulla urna, pharetra non tortor sed, sollicitudin molestie diam. Maecenas enim leo, feugiat eget vehicula id, sollicitudin vitae ante.";


app.get('/', cors(), function (req, res) {
    res.sendFile(path.join(__dirname, '/mock.html'));
});


function isEmpty(obj) {
    for (var prop in obj) {
        if (obj.hasOwnProperty(prop))
            return false;
    }
    return true;
}

app.post('/api/docx', cors(), async function (req, res) {
    console.log('docx------------>');
    let wellsfargo = req.body;

    const table = new Table({

        rows: [

            new TableRow({
                children: [

                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "DATE:",
                            style: "tableCell1",

                        })],
                        margins: {
                            top: convertInchesToTwip(0.40),
                            // bottom: convertInchesToTwip(0.20),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                        columnSpan: 3,
                    }),

                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.date,
                            style: "tableCell2",
                            // indent: {
                            //     left: 1000,
                            // },
                        })],
                        margins: {
                            top: convertInchesToTwip(0.40),
                            // bottom: convertInchesToTwip(0.20),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                        columnSpan: 3,
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "PROJECT:",
                            style: "tableCell1",
                            // indent: {
                            //     left: 1000,
                            // },
                        })],
                        margins: {
                            top: convertInchesToTwip(0.40),
                            // bottom: convertInchesToTwip(0.20),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                        columnSpan: 3,
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },

                        children: [new Paragraph({
                            text: wellsfargo.projectName,
                            style: "tableCell2",
                            // indent: {
                            //     left: 1000,
                            // },
                        })],
                        margins: {
                            top: convertInchesToTwip(0.40),
                            // bottom: convertInchesToTwip(0.20),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                        columnSpan: 3,
                    }),

                ],
            }),
            new TableRow({
                children: [

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "FROM(Consultant):",
                            style: "tableCell1",

                        })],
                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [
                            new Paragraph({
                                text: "Lynxspring, Inc.",
                                style: "tableCell1",
                            }),
                            new Paragraph({
                                text: "1210 NE Windsor",
                                style: "tableCell1",
                            }),
                            new Paragraph({
                                text: "Lee’s Summit, Missouri 64086",
                                style: "tableCell1",
                            }),
                        ],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "TO(Client):",
                            style: "tableCell1",
                            // indent: {
                            //     left: 1000,
                            // },
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },

                        children: [
                            new Paragraph({
                                text: wellsfargo.clientName,
                                style: "tableCell2",
                            }),
                            new Paragraph({
                                text: wellsfargo.address2,
                                style: "tableCell2",
                            }),
                            new Paragraph({
                                text: wellsfargo.cityState,
                                style: "tableCell2",
                            }),
                        ],
                        columnSpan: 3,
                    }),

                ],
            }),

            new TableRow({
                children: [

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "SENT VIA:",
                            style: "tableCell1",

                        })],
                        margins: {
                            top: convertInchesToTwip(0.20),
                            // bottom: convertInchesToTwip(0.20),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.email,
                            color: "FF0000",
                            style: "tableCell2",
                            // indent: {
                            //     left: 500,
                            // },
                        })],
                        margins: {
                            top: convertInchesToTwip(0.20),
                            // bottom: convertInchesToTwip(0.20),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                        columnSpan: 9,
                    }),


                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.SINGLE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.SINGLE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "SERVICES:",
                            style: "tableCell1",

                        })],
                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.SINGLE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.SINGLE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.services,
                            color: "FF0000",
                            style: "tableCell2",
                            // indent: {
                            //     left: 500,
                            // },
                        })],
                        columnSpan: 9,
                    }),
                ],

            }),

            new TableRow({
                children: [
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "PROJECT_DESCRIPTION:",
                            style: "tableCell1",

                        })],

                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.projectDescription,
                            color: "FF0000",
                            style: "tableCell2",
                            // indent: {
                            //     left: 500,
                            // },
                        })],
                        columnSpan: 9,

                    }),
                ],

            }),
            new TableRow({

                children: [
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "TYPE OF PROPOSAL:",
                            style: "tableCell1",

                        })],
                        margins: {
                            top: convertInchesToTwip(0.30),
                            // bottom: convertInchesToTwip(0.69),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.hourlySum,
                            color: "FF0000",
                            style: "tableCell2",
                            indent: {

                                left: 500,
                            },
                        })],
                        margins: {
                            top: convertInchesToTwip(0.30),
                            // bottom: convertInchesToTwip(0.69),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                        columnSpan: 9,

                    }),

                ],


            }),
            new TableRow({

                children: [
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "Void if not signed in 60 days.  Reimbursable expenses @ 1.15 times cost.  Additional Services are hourly per “Hourly Rate Schedule” provided in Part 4.2",
                            style: "tab",

                        })],
                        margins: {
                            // top: convertInchesToTwip(0.30),
                            bottom: convertInchesToTwip(0.20),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                        columnSpan: 12,
                    }),

                ],


            }),
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.SINGLE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.SINGLE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "LIABILITY LIMITATION:",
                            style: "tableCell1",

                        })],
                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.SINGLE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.SINGLE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "Consultant's Compensation (Refer to “PART 1.5”)",
                            color: "FF0000",
                            style: "tab1",
                            // indent: {
                            //     left: 500,
                            // },
                        })],
                        columnSpan: 9,
                    }),
                ],

            }),
            new TableRow({

                children: [
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "PROPOSED SERVICES & FEES:",
                            style: "tableCell1",

                        })],

                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "See Scope of Services Summary on Page 2 & SOW on Page 5",
                            style: "tab1",
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 9,

                    }),

                ],


            }),
            new TableRow({

                children: [
                    new TableCell({

                        children: [new Paragraph({
                            text: "Description",
                            alignment: AlignmentType.CENTER,
                            // numbering: {
                            //     reference: "ref1",
                            //     level: 0,
                            // },
                            style: "tableCell3",

                        })],

                        columnSpan: 5,
                    }),


                    new TableCell({

                        children: [new Paragraph({
                            text: "Hourly rate(refer to part 4.2)",
                            style: "tableCell3",
                            alignment: AlignmentType.CENTER,
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 7,

                    }),

                ],


            }),
            new TableRow({

                children: [
                    new TableCell({

                        children: [new Paragraph({
                            text: "",
                            numbering: {
                                reference: "ref1",
                                level: 0,
                            },
                            style: "tableCell4",

                        })],

                        columnSpan: 5,
                    }),


                    new TableCell({

                        children: [new Paragraph({
                            text: "",
                            // numbering: {
                            //     reference: "ref1",
                            //     level: 0,
                            // },
                            style: "tableCell4",
                            // alignment: AlignmentType.RIGHT,
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 7,

                    }),

                ],


            }),
            new TableRow({

                children: [
                    new TableCell({

                        children: [new Paragraph({
                            text: "",
                            numbering: {
                                reference: "ref1",
                                level: 0,
                            },
                            style: "tableCell4",

                        })],

                        columnSpan: 5,
                    }),


                    new TableCell({

                        children: [new Paragraph({
                            text: "",
                            // numbering: {
                            //     reference: "ref1",
                            //     level: 0,
                            // },
                            style: "tableCell4",
                            // alignment: AlignmentType.RIGHT,
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 7,

                    }),

                ],


            }),
            new TableRow({

                children: [
                    new TableCell({

                        children: [new Paragraph({
                            text: "",
                            numbering: {
                                reference: "ref1",
                                level: 0,
                            },
                            style: "tableCell4",

                        })],

                        columnSpan: 5,
                    }),


                    new TableCell({

                        children: [new Paragraph({
                            text: "",
                            // numbering: {
                            //     reference: "ref1",
                            //     level: 0,
                            // },
                            style: "tableCell4",
                            // alignment: AlignmentType.RIGHT,
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 7,

                    }),

                ],


            }),
            new TableRow({

                children: [
                    new TableCell({

                        children: [new Paragraph({
                            text: "",
                            numbering: {
                                reference: "ref1",
                                level: 0,
                            },
                            style: "tableCell4",

                        })],

                        columnSpan: 5,
                    }),


                    new TableCell({

                        children: [new Paragraph({
                            text: "",
                            // numbering: {
                            //     reference: "ref1",
                            //     level: 0,
                            // },
                            style: "tableCell4",
                            // alignment: AlignmentType.RIGHT,
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 7,

                    }),

                ],


            }),
            new TableRow({
                // borders: {
                //     top: {
                //         style: BorderStyle.SINGLE,
                //         size: 10,
                //         // color: "000000",
                //     },
                //     bottom: {
                //         style: BorderStyle.SINGLE,
                //         size: 1,
                //         // color: "000000",
                //     },
                //     left: {
                //         style: BorderStyle.SINGLE,
                //         size: 10,
                //         // color: "ff0000",
                //     },
                //     right: {
                //         style: BorderStyle.SINGLE,
                //         size: 10,
                //         // color: "ff0000",
                //     },
                // },

                children: [
                    new TableCell({


                        children: [new Paragraph({
                            text: "",
                            numbering: {
                                reference: "ref1",
                                level: 0,
                            },
                            style: "tableCell4",

                        })],

                        columnSpan: 5,
                    }),


                    new TableCell({

                        children: [new Paragraph({
                            text: "",

                            // numbering: {
                            //     reference: "ref1",
                            //     level: 0,
                            // },
                            style: "tableCell4",
                            // alignment: AlignmentType.RIGHT,
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 7,

                    }),

                ],


            }),
        ]
    });
    // ------------------------table-1-end-----------------



    // ------------------------table-2-end-----------------
    const table1 = new Table({
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.SINGLE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "PAYMENT TERMS: Net 10 days – 1.5% per month service charge over 30 days.",
                            style: "tableCell1",

                        })],
                        columnSpan: 6,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.SINGLE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "INVOICE SCHEDULE: See SOW on Page 5",
                            style: "tableCell1",

                        })],
                        columnSpan: 6,
                    }),


                ],

            }),
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.SINGLE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.SINGLE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "THIS EXECUTED PROPOSAL FORM represents the Agreement between the Consultant and the Client and supersedes all prior negotiations, representations or agreements, either written or oral. This Proposal may be amended only by written instrument signed by both parties. Upon execution and receipt of this form, together with any retainer amounts indicated, the Consultant will proceed with the project. The person signing acceptance below for the Client does hereby certify that he or she is fully authorized and empowered to execute this Instrument, and to bind the Client hereto, and does in fact so execute this Instrument.",
                            style: "tab",

                        })],
                        columnSpan: 12,
                    }),


                ],

            }),
            new TableRow({

                children: [
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "PROPOSED FOR THE CONSULTANT:",
                            style: "tableCell1",

                        })],

                        columnSpan: 6,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "ACCEPTED FOR THE CLIENT:",
                            style: "tableCell1",
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 6,

                    }),

                ],


            }),

            new TableRow({

                children: [
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "LYNXSPRING, INC.",
                            style: "tableCell11",


                        })],
                        // margins: {
                        //     top: convertInchesToTwip(-0.60),
                        //     // bottom: convertInchesToTwip(0.60),
                        //     // left: convertInchesToTwip(0.69),
                        //     // right: convertInchesToTwip(0.69),
                        // },

                        columnSpan: 6,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.clientName,
                            style: "tableCell12",
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 6,

                    }),

                ],


            }),
            new TableRow({

                children: [
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "By:",
                            style: "tableCell1",

                        })],
                        margins: {
                            top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.60),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },

                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.proposedBy,
                            style: "tableCell2",
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 3,

                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "By :",
                            style: "tableCell1",

                        })],
                        margins: {
                            top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.60),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },

                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.acceptedBy,
                            style: "tableCell2",
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 3,

                    }),

                ],


            }),
            new TableRow({

                children: [
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "Name (printed):",
                            style: "tableCell1",

                        })],

                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.proposedName,
                            style: "tableCell2",
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 3,

                    }),

                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "Name (printed):",
                            style: "tableCell1",

                        })],

                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.acceptedName,
                            style: "tableCell2",
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 3,

                    }),

                ],


            }),
            new TableRow({

                children: [
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "Title:",
                            style: "tableCell1",

                        })],

                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.proposedTitle,
                            style: "tableCell2",
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 3,

                    }),

                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "Title:",
                            style: "tableCell1",

                        })],

                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.acceptedTitle,
                            style: "tableCell2",
                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 3,

                    }),

                ],


            }),
            new TableRow({

                children: [
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "Date:",
                            style: "tableCell1",

                        })],

                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.proposedDate,
                            style: "tableCell2",

                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 3,

                    }),

                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: "Date Signed:",
                            style: "tableCell1",

                        })],

                        columnSpan: 3,
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                                color: "000000",
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                                color: "000000",
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                                color: "ff0000",
                            },
                        },
                        children: [new Paragraph({
                            text: wellsfargo.acceptedDateSigned,
                            style: "tableCell2",

                            // indent: {
                            //     left: 500,
                            // },
                        })],

                        columnSpan: 3,

                    }),

                ],


            }),



        ],
    });
    // SECOND PAGE

    const tablebox = new Table({
        rows: [
            new TableRow({

                children: [
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Discovery Phase Services", style: "normalPara1", })],
                        columnSpan: 4,
                        margins: {
                            top: convertInchesToTwip(0.30),
                            // bottom: convertInchesToTwip(0.20),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Bidding & Negotiation Phase Services", style: "normalPara1", })],
                        columnSpan: 4,
                        margins: {
                            top: convertInchesToTwip(0.30),
                            // bottom: convertInchesToTwip(0.20),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Studies", style: "normalPara1", })],
                        columnSpan: 5,
                        margins: {
                            top: convertInchesToTwip(0.30),
                            // bottom: convertInchesToTwip(0.20),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),

                ]

            }),

            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Initial investigation with report",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Attend pre–bid meeting",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "HVAC System’s",
                            style: "normalPara",
                        })],
                        columnSpan: 5,
                    }),
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Prepare Functional Specification documents",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Respond to bidder questions",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Lighting System's",
                            style: "normalPara",
                        })],
                        columnSpan: 5,
                    }),
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Preliminary design documents",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Attend bid opening meeting",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Power System's",
                            style: "normalPara",
                        })],
                        columnSpan: 5,
                    }),
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Preliminary Bill-of-Materials",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Analysis of bids",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Building Energy Analysis",
                            style: "normalPara",
                        })],
                        columnSpan: 5,
                    }),
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Estimate probable cost",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Value Engineering",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Economic Analysis ",
                            style: "normalPara",
                        })],
                        columnSpan: 5,
                    }),
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Attend discovery review meetings ",
                            style: "normalPara",
                        })],
                        columnSpan: 6,

                    }),


                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Facility Audit's",
                            style: "normalPara",
                        })],
                        columnSpan: 5,
                    }),
                ]
            }),
            new TableRow({
                children: [

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "",
                            style: "normalPara",
                        })],
                        columnSpan: 8,

                    }),

                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Feasibility Study",
                            style: "normalPara",
                        })],
                        columnSpan: 5,
                    }),
                ]
            }),
            new TableRow({
                children: [

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "",
                            style: "normalPara",
                        })],
                        columnSpan: 8,

                    }),

                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "IAQ Investigation",
                            style: "normalPara",
                        })],
                        columnSpan: 5,
                    }),
                ]
            }),
            new TableRow({

                children: [
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Development Phase Services", style: "normalPara1", })],
                        columnSpan: 4,
                        // margins: {
                        //     top: convertInchesToTwip(0.90),
                        //     // bottom: convertInchesToTwip(0.20),
                        //     // left: convertInchesToTwip(0.69),
                        //     // right: convertInchesToTwip(0.69),
                        // },
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Specifications", style: "normalPara1", })],
                        columnSpan: 4,
                        // margins: {
                        //     top: convertInchesToTwip(0.90),
                        //     // bottom: convertInchesToTwip(0.20),
                        //     // left: convertInchesToTwip(0.69),
                        //     // right: convertInchesToTwip(0.69),
                        // },
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Electrical Coordination Study",
                            style: "normalPara",
                        })],
                        columnSpan: 5,
                    }),

                ]

            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Development of Hardware Platform Design",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Provide full book specification",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "",
                            style: "normalPara",
                        })],
                        columnSpan: 5,
                    }),
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Development of Security Platform",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Provide Marketing Documentation",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "",
                            style: "normalPara",
                        })],
                        columnSpan: 5,
                    }),
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Development of HMI Appliance",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "RED MARK” specifications for others",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "",
                            style: "normalPara",
                        })],
                        columnSpan: 5,
                    }),
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Development of Device Driver",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    // new TableCell({
                    //     width: {
                    //         size: 3,
                    //         type: WidthType.PERCENTAGE,
                    //     },
                    //     borders: {
                    //         top: {
                    //             style: BorderStyle.NONE,
                    //             size: 15,
                    //         },
                    //         bottom: {
                    //             style: BorderStyle.NONE,
                    //             size: 16,
                    //         },
                    //         left: {
                    //             style: BorderStyle.NONE,
                    //             size: 1,
                    //         },
                    //         right: {
                    //             style: BorderStyle.NONE,
                    //             size: 1,
                    //         },
                    //     },
                    //     children: [new Paragraph("")],
                    //     // columnSpan: 2,
                    // }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Special Conditions", style: "normalPara1", })],
                        columnSpan: 5,
                        // margins: {
                        //     top: convertInchesToTwip(0.90),
                        //     // bottom: convertInchesToTwip(0.20),
                        //     // left: convertInchesToTwip(0.69),
                        //     // right: convertInchesToTwip(0.69),
                        // },
                    }),
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Software Custom Module Development",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Overtime authorization at higher hourly ",
                            style: "normalPara",
                        })],
                        columnSpan: 5,

                    }),


                ]
            }),

            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Development of Documentation",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Rate of 1.5 times Hourly Rate Schedule", style: "normalPara", })],
                        columnSpan: 5,

                    }),
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Development of Web Appliance",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    // new TableCell({
                    //     width: {
                    //         size: 3,
                    //         type: WidthType.PERCENTAGE,
                    //     },
                    //     borders: {
                    //         top: {
                    //             style: BorderStyle.NONE,
                    //             size: 15,
                    //         },
                    //         bottom: {
                    //             style: BorderStyle.NONE,
                    //             size: 1,
                    //         },
                    //         left: {
                    //             style: BorderStyle.NONE,
                    //             size: 1,
                    //         },
                    //         right: {
                    //             style: BorderStyle.NONE,
                    //             size: 1,
                    //         },
                    //     },
                    //     children: [new Paragraph("")],
                    //     // columnSpan: 2,
                    // }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Coordination Services", style: "normalPara1", })],
                        columnSpan: 4,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 5,

                    }),
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Development of Automated Processes",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Coordinate with Client",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "",
                            style: "normalPara",
                        })],
                        columnSpan: 5,

                    }),


                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Others special systems:  Describe",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Coordinate with Client Customer",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    // new TableCell({
                    //     width: {
                    //         size: 3,
                    //         type: WidthType.PERCENTAGE,
                    //     },

                    //     children: [new Paragraph("")],
                    //     // columnSpan: 2,
                    // }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Expenses", style: "normalPara1", })],
                        columnSpan: 5,

                    }),


                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Estimate probable cost",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Coordinate with Agency Officials",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "R = Reimbursable",
                            style: "normalPara",
                        })],
                        columnSpan: 5,

                    }),


                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Attend development review meetings ",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "Other Consultants:  Describe",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "NR = Non-reimbursable",
                            style: "normalPara",
                        })],
                        columnSpan: 5,

                    }),


                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph("")],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "",
                            style: "normalPara",
                        })],
                        columnSpan: 2,

                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "",
                            style: "normalPara",
                        })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },

                        children: [new Paragraph("")],
                        // columnSpan: 2,
                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({
                            text: "RAM = Reimbursable Against Maximum",
                            style: "normalPara",
                        })],
                        columnSpan: 5,

                    }),


                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara1", })],
                        columnSpan: 4,

                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara1", })],

                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,

                    }),
                    new TableCell({

                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara12", })],

                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 5,

                    }),

                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Deployment Phase Services", style: "normalPara1", })],
                        columnSpan: 4,

                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara1", })],

                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,

                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },

                        children: [new Paragraph({ text: "R", style: "normalPara12", })],

                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Delivery/Freight/Postage", style: "normalPara", })],
                        columnSpan: 5,

                    }),

                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({})],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Attend pre-deployment meeting", style: "normalPara", })],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({ text: "NR", style: "normalPara12", })],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Long Distance Telephone", style: "normalPara", })],
                        columnSpan: 5,
                    }),

                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({})],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Attend deployment progress meetings", style: "normalPara", })],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({ text: "R", style: "normalPara12", })],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Mileage/Parking/Tolls", style: "normalPara", })],
                        columnSpan: 5,
                    }),

                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({})],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Review marketing collateral", style: "normalPara", })],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({ text: "R", style: "normalPara12", })],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "All job related printing & reproductions", style: "normalPara", })],
                        columnSpan: 5,
                    }),

                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({})],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Perform training", style: "normalPara", })],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 5,
                    }),

                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({})],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Respond to RFI's", style: "normalPara", })],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 5,
                    }),

                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({})],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Manage/Assist Beta Roll-Out", style: "normalPara", })],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,
                    }),
                    // new TableCell({
                    //     borders: {
                    //         top: {
                    //             style: BorderStyle.NONE,
                    //             size: 15,
                    //         },
                    //         bottom: {
                    //             style: BorderStyle.NONE,
                    //             size: 16,
                    //         },
                    //         left: {
                    //             style: BorderStyle.NONE,
                    //             size: 1,
                    //         },
                    //         right: {
                    //             style: BorderStyle.NONE,
                    //             size: 1,
                    //         },
                    //     },
                    //     children: [new Paragraph({})],
                    // }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Legend", style: "normalPara1", })],
                        columnSpan: 5,
                    }),

                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({})],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Attend deployment progress meetings", style: "normalPara", })],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({ text: "R", style: "normalPara12", })],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Mileage/Parking/Tolls", style: "normalPara", })],
                        columnSpan: 5,
                    }),

                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({})],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Prepare O & M manuals", style: "normalPara", })],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({ text: "X", style: "normalPara12", })],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Included Services", style: "normalPara", })],
                        columnSpan: 5,
                    }),

                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({})],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Perform building commissioning and validation", style: "normalPara", })],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({ text: "O", style: "normalPara12", })],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Optional Line Items Services", style: "normalPara", })],
                        columnSpan: 5,
                    }),

                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({})],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Perform substantial completion inspection", style: "normalPara", })],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({ text: "H", style: "normalPara12", })],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Hourly Services", style: "normalPara", })],
                        columnSpan: 5,
                    }),

                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({ text: "P +", style: "normalPara12", })],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Per Diem + expenses", style: "normalPara", })],
                        columnSpan: 5,
                    }),

                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 2,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({})],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "", style: "normalPara", })],
                        columnSpan: 3,
                    }),
                    new TableCell({
                        width: {
                            size: 3,
                            type: WidthType.PERCENTAGE,
                        },
                        children: [new Paragraph({ text: "C +", style: "normalPara12", })],
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 15,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 16,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "(Cost + Overhead) times multiplier", style: "normalPara", })],
                        columnSpan: 5,
                    }),

                ],
            }),

        ],
        width: {
            size: 100,
            type: WidthType.PERCENTAGE,
        },
    })



    const tableadmin = new Table({
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Category", style: "normalPara5", })],
                        columnSpan: 2,
                        margins: {
                            // top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.20),
                            left: convertInchesToTwip(0.50),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Rate", style: "normalPara5", })],
                        columnSpan: 2,
                        margins: {
                            // top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.20),
                            // left: convertInchesToTwip(0.69),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Administrative", style: "normalPara6", })],
                        columnSpan: 2,
                        margins: {
                            top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.20),
                            left: convertInchesToTwip(0.50),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "$125.00", style: "normalPara6", })],
                        columnSpan: 2,
                        margins: {
                            top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.20),
                            // left: convertInchesToTwip(0.50),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Technician", style: "normalPara6", })],
                        columnSpan: 2,
                        margins: {
                            // top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.20),
                            left: convertInchesToTwip(0.50),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "$150.00", style: "normalPara6", })],
                        columnSpan: 2,
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Trainer", style: "normalPara6", })],
                        columnSpan: 2,
                        margins: {
                            // top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.20),
                            left: convertInchesToTwip(0.50),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "$200.00", style: "normalPara6", })],
                        columnSpan: 2,

                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Consultant", style: "normalPara6", })],
                        columnSpan: 2,
                        margins: {
                            // top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.20),
                            left: convertInchesToTwip(0.50),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "$200.00", style: "normalPara6", })],
                        columnSpan: 2,
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Developer", style: "normalPara6", })],
                        columnSpan: 2,
                        margins: {
                            // top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.20),
                            left: convertInchesToTwip(0.50),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "$250.00", style: "normalPara6", })],
                        columnSpan: 2,
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Engineer", style: "normalPara6", })],
                        columnSpan: 2,
                        margins: {
                            // top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.20),
                            left: convertInchesToTwip(0.50),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "$250.00", style: "normalPara6", })],
                        columnSpan: 2,
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Project Management", style: "normalPara6", })],
                        columnSpan: 2,
                        margins: {
                            // top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.20),
                            left: convertInchesToTwip(0.50),
                            // right: convertInchesToTwip(0.69),
                        },
                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "$300.00", style: "normalPara6", })],
                        columnSpan: 2,
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Executive", style: "normalPara6", })],
                        columnSpan: 2,
                        margins: {
                            // top: convertInchesToTwip(0.10),
                            // bottom: convertInchesToTwip(0.20),
                            left: convertInchesToTwip(0.50),
                            // right: convertInchesToTwip(0.69),
                        },

                    }),

                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "$500.00", style: "normalPara6", })],
                        columnSpan: 2,
                    }),
                ],
            }),
        ],
        width: {
            size: 50,
            type: WidthType.PERCENTAGE,
        },

    });


    const lasttable = new Table({
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        children: [new Paragraph({}), new Paragraph({})],
                        verticalAlign: VerticalAlign.CENTER,
                        width: {
                            size: 15,
                            type: WidthType.PERCENTAGE,
                        },
                    }),
                    new TableCell({
                        children: [new Paragraph({ text: "Description of Deliverables", style: "normalPara7", alignment: AlignmentType.CENTER }), new Paragraph({})],
                        textDirection: TextDirection.CENTER,
                        width: {
                            size: 30,
                            type: WidthType.PERCENTAGE,
                        },
                    }),
                    new TableCell({
                        children: [new Paragraph({ text: "Estimated Start & End Dates ", style: "normalPara7", alignment: AlignmentType.CENTER }), new Paragraph({})],
                        textDirection: TextDirection.CENTER,
                        width: {
                            size: 35,
                            type: WidthType.PERCENTAGE,
                        },
                    }),
                    new TableCell({
                        children: [new Paragraph({ text: "Amount", style: "normalPara7", alignment: AlignmentType.CENTER }), new Paragraph({})],
                        textDirection: TextDirection.CENTER,
                        width: {
                            size: 20,
                            type: WidthType.PERCENTAGE,
                        },
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Milestone 1",
                                style: "normalPara7",
                                alignment: AlignmentType.LEFT,
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                        ],
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Mobilization",
                                style: "normalPara11",
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Est. Start Date:",
                                style: "normalPara9",
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                            new Paragraph({
                                text: "Upon receipt of Purchase Order",
                                style: "normalPara11",
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Due Upon execution of Proposal 25% Labor",
                                style: "normalPara11",
                            }),
                            new Paragraph({
                                text: wellsfargo.amount1,
                                style: "aside",
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Milestone 2",
                                style: "normalPara7",
                                alignment: AlignmentType.LEFT,
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                        ],
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Submittal drawings",
                                style: "normalPara11",
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Start Date:",
                                style: "normalPara9",
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                            new Paragraph({
                                text: "Upon approval of submittals",
                                style: "normalPara11",
                            }),
                            new Paragraph({
                                text: "End Date:",
                                style: "normalPara9",
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                            new Paragraph({
                                text: "Upon completion of controller programming.  8-12 Weeks depending on panel requirements",
                                style: "normalPara11",
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "100% material + 25% labor",
                                style: "normalPara11",
                            }),
                            new Paragraph({
                                text: wellsfargo.amount2,
                                style: "aside",
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Milestone 3",
                                style: "normalPara7",
                                alignment: AlignmentType.LEFT,
                                indent: { left: convertInchesToTwip(0.1) },

                            }),
                        ],
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Substantial Completion:",
                                style: "normalPara11",
                            }),
                            new Paragraph({
                                text: "• Material Shipment",
                                style: "normalPara11",
                            }),
                            new Paragraph({
                                text: "• Site Commissioning",
                                style: "normalPara11",
                            }),
                            new Paragraph({
                                text: "• BAS System running and proving control. ",
                                style: "normalPara11",
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Start Date:",
                                style: "normalPara9",
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                            new Paragraph({
                                text: "Upon shipment of materials",
                                style: "normalPara11",
                            }),
                            new Paragraph({
                                text: "End Date: ",
                                style: "normalPara9",
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                            new Paragraph({
                                text: "Site Commission reasonably completed, and system is in operation. ",
                                style: "normalPara11",
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "45% Labor",
                                style: "normalPara11",
                            }),
                            new Paragraph({
                                text: wellsfargo.amount3,
                                style: "aside",
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                ],
            }),

            new TableRow({
                children: [
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Milestone 4",
                                style: "normalPara7",
                                alignment: AlignmentType.LEFT,
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                        ],
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Closeout & Final Acceptance",
                                style: "normalPara11",
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Start Date:",
                                style: "normalPara9",
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                            new Paragraph({
                                text: "Site Commission reasonably completed.",
                                style: "normalPara11",
                            }),
                            new Paragraph({
                                text: "End Date:",
                                style: "normalPara9",
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                            new Paragraph({
                                text: "4 weeks after start of Site Commission",
                                style: "normalPara11",
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                text: "Remaining 5% Labor",
                                style: "normalPara11",
                            }),
                            new Paragraph({
                                text: wellsfargo.amount4,
                                style: "aside",
                                indent: { left: convertInchesToTwip(0.1) },
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                ],
            }),
        ],
        width: {
            size: 100,
            type: WidthType.PERCENTAGE,
        },
    });

    const total = new Table({
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [
                            new Paragraph({
                                text: "",
                                style: "normalPara2",
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [
                            new Paragraph({
                                text: "",
                                style: "normalPara2",
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [
                            new Paragraph({
                                text: "Total:",
                                style: "normalPara10",
                                alignment: AlignmentType.RIGHT
                            }),
                        ],
                        verticalAlign: VerticalAlign.LEFT,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [
                            new Paragraph({
                                text: wellsfargo.totallast,
                                style: "aside",
                                alignment: AlignmentType.CENTER
                            }),
                        ],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                ],
            }),

        ],
        width: {
            size: 100,
            type: WidthType.PERCENTAGE,
        },
    });


    const footer = new Table({
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "Lynxspring Proposal for Services Agreement", style: "footer", })],
                        verticalAlign: VerticalAlign.CENTER,
                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({

                            children: [

                                new TextRun({
                                    children: ["Page", PageNumber.CURRENT, " Of ", PageNumber.TOTAL_PAGES,],
                                    italics: true,
                                    font: {
                                        name: "Calibri",
                                    },
                                    margins: {
                                        // top: convertInchesToTwip(0.30),
                                        // bottom: convertInchesToTwip(0.20),
                                        left: convertInchesToTwip(0.69),
                                        // right: convertInchesToTwip(0.69),
                                    },
                                }),
                            ],

                        })],

                    }),
                    new TableCell({
                        borders: {
                            top: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            bottom: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            left: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                            right: {
                                style: BorderStyle.NONE,
                                size: 1,
                            },
                        },
                        children: [new Paragraph({ text: "April 25, 2022", style: "footer", })],
                        textDirection: TextDirection.CENTER,
                    }),
                ],
            }),

        ],
        width: {
            size: 100,
            type: WidthType.PERCENTAGE,
        },
    });

    const doc = new Document({
        numbering: {
            config: [{
                reference: 'ref1',
                levels: [{
                    level: 0,
                    format: LevelFormat.DECIMAL,
                    text: '%1.',
                    start: 1,
                }],
            }]
        },
        styles: {
            default: {
                heading1: {
                    run: {
                        font: "Calibri",
                        size: 36,
                        bold: true,
                        color: "005CA9",

                    },
                    paragraph: {
                        alignment: AlignmentType.CENTER,
                        spacing: { line: 300, before: 1 * 72 * 0.1, after: 1 * 72 * 0.05 },

                    },

                },
                heading2: {
                    run: {
                        font: "Calibri",
                        size: 22,
                        bold: true,
                    },
                    paragraph: {
                        spacing: { line: 340 },
                        // indent: {
                        //     left: 720,
                        // },
                    },
                },
            },

            paragraphStyles: [{
                id: "headerImage",
                name: "Header Image",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                paragraph: {
                    alignment: AlignmentType.CENTER,
                    spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                },
            },
            {
                id: "tableCell1",
                name: "Table Cell1",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 20,
                    bold: true,

                },
                paragraph: {
                    spacing: { line: 256, before: 10 * 72 * 0.1, after: 10 * 72 * 0.05 },
                    rightTabStop: TabStopPosition.MAX,
                    leftTabStop: 453.543307087,
                },
            },
            {
                id: "tableCell11",
                name: "Table Cell1",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 20,
                    bold: true,

                },
                // paragraph: {
                //     spacing: { line: 256, before: 10 * 72 * 0.1, after: 10 * 72 * 0.05 },
                //     rightTabStop: TabStopPosition.MAX,
                //     leftTabStop: 453.543307087,
                // },
            },
            {
                id: "tableCell12",
                name: "Table Cell1",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 20,
                    bold: true,
                    color: '0000ff'
                },
                // paragraph: {
                //     spacing: { line: 256, before: 10 * 72 * 0.1, after: 10 * 72 * 0.05 },
                //     rightTabStop: TabStopPosition.MAX,
                //     leftTabStop: 453.543307087,
                // },
            },
            {
                id: "tableCell2",
                name: "Table Cell2",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 20,
                    bold: true,
                    color: '0000ff'
                },

                paragraph: {
                    spacing: { line: 256, before: 10 * 72 * 0.1, after: 10 * 72 * 0.05 },
                    rightTabStop: TabStopPosition.MAX,
                    leftTabStop: 453.543307087,
                },
            },
            {
                id: "tableCell3",
                name: "Table Cell3",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 24,
                    bold: true,
                    // color: '0000ff'
                },

                numbering: {
                    reference: "ref1",
                    level: 0,
                },
                paragraph: {
                    spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                    rightTabStop: TabStopPosition.MAX,
                    leftTabStop: 453.543307087,
                },
            },
            {
                id: "tableCell4",
                name: "Table Cell4",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 20,
                    bold: true,
                    color: '0000ff',
                    // underline: {
                    //     type: UnderlineType.SINGLE,
                    //     color: "000000",
                    // },
                    wrap: {
                        type: TextWrappingType.SQUARE,
                        side: TextWrappingSide.BOTH_SIDES,
                    },
                },

                numbering: {
                    reference: "ref1",
                    level: 0,
                },
                paragraph: {
                    spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                    rightTabStop: TabStopPosition.MAX,
                    leftTabStop: 453.543307087,
                },
            },
            {
                id: "tab",
                name: "Tab",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 16,
                    // bold: true,
                    // color: '0000ff'
                },
                paragraph: {
                    // spacing: { line: 10, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                    rightTabStop: TabStopPosition.MAX,
                    leftTabStop: 453.543307087,
                },
            },
            {
                id: "normalPara12",
                name: "Normal Para",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 16,
                    bold: true,

                },

                paragraph: {
                    alignment: AlignmentType.CENTER
                },
            },
            {
                id: "tab1",
                name: "Tab1",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 20,
                    // bold: true,

                },
                paragraph: {
                    spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                    rightTabStop: TabStopPosition.MAX,
                    leftTabStop: 453.543307087,
                },
            },
            {
                id: "normalPara",
                name: "Normal Para",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 16,
                },
                // paragraph: {
                //     spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                //     rightTabStop: TabStopPosition.MAX,
                //     leftTabStop: 453.543307087,
                // },
            },

            {
                id: "normalPara6",
                name: "Normal Para",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 18,
                    margins: {
                        top: convertInchesToTwip(0.30),
                        // bottom: convertInchesToTwip(0.20),
                        // left: convertInchesToTwip(0.69),
                        // right: convertInchesToTwip(0.69),
                    },
                },
                // paragraph: {
                //     spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                //     rightTabStop: TabStopPosition.MAX,
                //     leftTabStop: 453.543307087,
                // },
            },
            {
                id: "normalPara1",
                name: "Normal Para",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 16,
                    bold: true,
                    underline: {
                        type: UnderlineType.SINGLE,
                        color: "000000",
                    },
                },
                // paragraph: {
                //     spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                //     rightTabStop: TabStopPosition.MAX,
                //     leftTabStop: 453.543307087,
                // },
            },

            {
                id: "normalPara10",
                name: "Normal Para",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 22,
                    bold: true,
                },

                // paragraph: {
                //     spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                //     rightTabStop: TabStopPosition.MAX,
                //     leftTabStop: 453.543307087,
                // },
            },
            {
                id: "normalPara5",
                name: "Normal Para",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 18,
                    bold: true,
                    underline: {
                        type: UnderlineType.SINGLE,
                        color: "000000",
                    },
                },

                // paragraph: {
                //     spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                //     rightTabStop: TabStopPosition.MAX,
                //     leftTabStop: 453.543307087,
                // },
            },
            {
                id: "normalPara9",
                name: "Normal Para",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 22,
                    bold: true,
                    underline: {
                        type: UnderlineType.SINGLE,
                        color: "000000",
                    },
                },

                // paragraph: {
                //     spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                //     rightTabStop: TabStopPosition.MAX,
                //     leftTabStop: 453.543307087,
                // },
            },

            {
                id: "normalPara7",
                name: "Normal Para",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 22,
                    bold: true,
                },
                paragraph: {
                    spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                    rightTabStop: TabStopPosition.MAX,
                    leftTabStop: 453.543307087,
                },
            },
            {
                id: "normalPara3",
                name: "Normal Para",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 18,
                },
                paragraph: {
                    spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                    rightTabStop: TabStopPosition.MAX,
                    leftTabStop: 453.543307087,
                },
            },
            {
                id: "normalPara2",
                name: "Normal Para2",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 20,

                },
                paragraph: {
                    alignment: AlignmentType.JUSTIFIED,
                },
            },
            {
                id: "normalPara11",
                name: "Normal Para2",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 22,

                },
                paragraph: {
                    alignment: AlignmentType.LEFT,
                    indent: { left: convertInchesToTwip(0.1) },
                },
            },
            {
                id: "normalPara4",
                name: "Normal Para",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: {
                    font: "Calibri",
                    size: 18,
                },
                paragraph: {
                    spacing: { line: 266, before: 30 * 72 * 0.1, after: 30 * 72 * 0.05 },
                    rightTabStop: TabStopPosition.MAX,
                    leftTabStop: 453.543307087,
                },
            },

            {
                id: "aside",
                name: "Aside",
                basedOn: "Normal",
                next: "Normal",
                run: {
                    font: "Calibri",
                    size: 22,
                    bold: true,
                    color: "005CA9",
                },
                // paragraph: {
                //     spacing: { line: 276 },
                //     indent: { left: convertInchesToTwip(0.5) },
                // },
            },
            {
                id: "wellSpaced",
                name: "Well Spaced",
                basedOn: "Normal",
                paragraph: {
                    spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },
                },
            },
            {
                id: "footer",
                name: "Numbered Para",
                basedOn: "Normal",
                next: "Normal",
                run: {
                    font: "Calibri",
                    size: 18,
                    italics: true,
                },
                // paragraph: {
                //     spacing: { line: 276, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 },

                // },
            },
            {
                id: "header1",
                name: "Numbered Para",
                basedOn: "Normal",
                next: "Normal",
                run: {
                    font: "Calibri",
                    size: 36,
                    bold: true,
                    color: "#707070"
                },
                // paragraph: {
                //     spacing: { line: 276, before: 20 * 72 * 0.1, after: 100 * 72 * 0.05 },

                // },
            },
            {
                id: "header2",
                name: "Numbered Para",
                basedOn: "Normal",
                next: "Normal",
                run: {
                    font: "Calibri",
                    size: 28,
                    bold: true,
                    color: "#707070"
                },
                paragraph: {
                    spacing: { line: 106, before: 10 * 72 * 0.1, after: 100 * 72 * 0.05 },

                },
            },
            {
                id: "header3",
                name: "Numbered Para",
                basedOn: "Normal",
                next: "Normal",
                run: {
                    font: "Calibri",
                    size: 36,
                    bold: true,
                    color: "#707070"
                },
                paragraph: {
                    spacing: { line: 276, before: 20 * 72 * 0.1, after: 100 * 72 * 0.05 },

                },
            },
            ],
        },
        // evenAndOddHeaderAndFooters: true,
        sections: [{
            properties: {
                titlePage: true,
                page: {
                    margin: {
                        top: 400,
                        right: 700,
                        bottom: 700,
                        left: 700,
                    },
                    size: {
                        orientation: PageOrientation.PORTRAIT,
                        // height: convertMillimetersToTwip(420),
                        width: convertMillimetersToTwip(240),
                    },
                },
            },
            // headers: {
            //     default: new Header({
            //         children: [
            //             new Paragraph({
            //                 text: "LYNXSPRING SCHEDULE OF INVOICES AND STATEMENT OF WORK",
            //                 style: "header3",
            //                 alignment: AlignmentType.CENTER,
            //                 verticalAlign: VerticalAlign.CENTER,
            //             }),
            //         ],
            //     }),
            //     even: new Header({
            //         children: [
            //             new Paragraph({
            //                 text: "LYNXSPRING SCOPE OF SERVICES SUMMARY",
            //                 style: "header1",
            //                 alignment: AlignmentType.CENTER
            //             }),
            //             new Paragraph({
            //                 text: "AGREEMENT FOR SERVICES",
            //                 style: "header2",
            //                 alignment: AlignmentType.CENTER
            //             }),
            //         ],
            //     }),
            // },
            footers: {
                default: new Footer({
                    children: [
                        footer,
                    ],
                }),
                even: new Footer({
                    children: [
                        footer,
                    ],
                }),
                first: new Footer({
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.LEFT,
                            children: [
                                footer,
                            ],
                        }),
                    ],
                }),
            },

            children: [

                new Paragraph({
                    children: [
                        new ImageRun({
                            data: fs.readFileSync("./images/lynxspring.png"),
                            transformation: {
                                width: 220,
                                height: 120,
                                // left: 200,
                            },
                            // alignment: AlignmentType.CENTER,
                        }),
                    ],
                    style: "headerImage",
                }),
                new Paragraph({
                    text: "PROPOSAL FOR SERVICES",
                    heading: HeadingLevel.HEADING_1,
                    alignment: AlignmentType.CENTER,


                }),

                table,
                table1,

                new Paragraph({
                    text: "LYNXSPRING SCOPE OF SERVICES SUMMARY",
                    style: "header1",
                    alignment: AlignmentType.CENTER,
                }),

                new Paragraph({
                    text: "AGREEMENT FOR SERVICES",
                    style: "header2",
                    alignment: AlignmentType.CENTER
                }),

                tablebox,


                new Paragraph({
                    spacing: {
                        after: 3000,
                        before: 2000,
                    },

                    children: [
                        new TextRun({
                            text: "Specific Inclusions / Exclusions:",
                            bold: true,
                            font: {
                                name: "Calibri",
                                size: 20,
                            },
                        }),
                        new TextRun({
                            text: " Stipulated Sum does not include cost of materials or expenses due to travel these and similar expenses will be considered reimbursable at 1.15 of actual cost and will be billed in addition to the Stipulated Sum quoted on Page 1 of 5 of this document.  Client shall pay all applicable sales, use, withholding and excise tax, and any other assessments.",
                            font: {
                                name: "Calibri",
                                size: 20,
                            },
                        }),
                    ],

                }),

                new Paragraph({
                    text: "LYNXSPRING GENERAL TERMS & CONDITIONS",
                    style: "header3",
                    alignment: AlignmentType.CENTER
                }),

                new Paragraph({
                    text: "PART 1 - GENERAL",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "1.1.	PROPOSAL CONFIDENTIALITY: The information in this proposal shall not be disclosed outside the Client’s organization and shall not be duplicated, used or disclosed in whole or in part for any purpose other than to evaluate the proposal.  If a contract is awarded to Lynxspring, Inc. as a result of or in connection with the submission of this proposal, the Client shall have the right to duplicate, use or disclose the information to the extent provided by the contract.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "1.2.	CREDITS: The Consultant shall have the right to include representations of the design of the Project, including photographs, among the Consultant’s promotional and professional materials, subject to that which if the Client has previously advised the Consultant in writing of the specific information considered by the Client to be confidential or proprietary. The Client will provide professional credit for the Consultant in promotional materials for the Project.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "1.3.	ORDERLY PROGRESS: Fees are based upon an orderly progression of work (by the Consultant) defined in the SOW with concurrent input and approvals (by the Client) from Discovery, through Development, and on to Deployment. Revisions to the Project’s design that are not consistent with previous decisions, or approvals on drawings previously received from Client, or that are made as a result of the Client’s, or other Client’s consultants, failure to make such decisions in a timely manner, are not included in any stipulated, lump-sum fees.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "1.4.	QUALIFICATION LIMITATION FOR FIXED FEES: Should the Scope of Work required for completion of the Project be different than that described in the proposal, or should the conditions under which the work is to be performed be changed through no fault of the Consultant, fee amounts may be amended by mutual agreement of both parties. Both parties prior to commencement of work shall sign a written agreement.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "1.5.	LIMITATION OF LIABILITY: Unless otherwise indicated in the proposal and only to the fullest extent permitted by law, Client agrees to limit the liability of the Consultant, its officers, shareholders, and employees, for any negligent acts, errors, omissions, or breaches of contract or warranty arising out of the performance of Consultant's services under this agreement, to the amount of  Consultant's Compensation  regardless of the number of claims or the number of parties prosecuting claims against Consultant.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "1.6.	TERMINATION: This Agreement may be terminated by either party upon not less than seven (7) days written notice should the other party fail substantiality to perform in accordance with the terms and Conditions contained herein, through no fault of the party initiating the termination.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "TERMINATION EXPENSES: In the event of suspension of services through no fault of the Consultant, the Consultant shall be compensated for services performed prior to the termination, together with all Reimbursable Expenses then due. The Client shall also reimburse the Consultant for all of the Consultant’s termination expenses including but not limited to, those associated with demobilization, re-assignment of personnel, and space and equipment costs.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "1.8.	SUBMITTALS AND APPROVALS:  Unless otherwise indicated, the Consultant will submit Documents to the Client for review, comment and approval.  The Consultant will revise and re-submit Documents, or provide modification documents as required by agencies having legal jurisdiction, in order to secure their approval, when such modifications are normal, reasonable, timely and customary.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "PART 2 - ADDITIONAL SERVICES",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "2.1.	Additional services shall be paid for by the Client based on hourly rates or stipulated sums indicated in the Proposal or Agreement, in addition to the compensation for the accepted Proposed Services & Fees.  Unless otherwise indicated within the Proposal, Agreement, or Scope of Services Description, the following services are not included in the Consultant’s Services fee.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "2.1.1.	Revisions to documents that are: inconsistent with instructions previously given by the Client, or due to changes required as a result of the Client’s or Client’s Consultant’s failure to render decisions in a timely manner.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "2.1.2. PROTO-TYPE RE-DESIGN due to subsequent changes by the client to the prototypical design criteria after Consultant has commenced with documents or products for a specific deployment.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "2.1.3. Billable Agency Review re-submittals: If the Consultant is asked to provide resubmittal because of the following circumstances, it will be considered a billable review submittal subject to additional costs to the client:  (1) changes that occurred to the documents or products after initial submittal that the Consultant was not aware of, or (2) information prompting the resubmittal was not part of publicly available documentation, or (3) review agency requires resubmittal of information that was contained in previous submittals, or (4) that are otherwise interpretive in nature. ",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                    spacing: {
                        after: 4000,
                    },
                }),

                new Paragraph({
                    text: "LYNXSPRING SCHEDULE OF INVOICES AND STATEMENT OF WORK",
                    style: "header3",
                    alignment: AlignmentType.CENTER
                }),

                new Paragraph({
                    text: "PART 3 - PAYMENTS TO THE CONSULTANT",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                    spacing: {
                        before: 500,
                    },
                }),

                new Paragraph({
                    text: "3.1.	INVOICES: Payments for invoices shall be made within 30 days of receipt or lesser time as is appropriate, for completion of services and for reimbursable expenses within that billing period. The Client shall not withhold payment to Consultant subject to claims or potential claims arising out of the construction for the work.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "3.2.	NOTIFICATION OF OBJECTIONS: If the Client objects to all or any portion of an invoice, the Client shall so notify the Consultant within seven (7) calendar days of receipt, identifying the cause for disagreement, and pay when due that portion of the invoice, if any, not in dispute. Disputes do not change the date payment is due, nor extend the grace period.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "3.3.	CHARGES FOR REIMBURSABLE EXPENSES are in addition to the amounts charged for Fees, and will be billed at a multiple of 1.15  the actual cost charged to the Consultant by outside service agencies, or at the specified rates for in-house expenses indicated.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "3.4.	LATE PAYMENT: A service charge of 1.5% per month of the adjusted previous balance or the maximum governmental allowed interest (or a minimum service charge of $20.00 per invoice) may be added to the Client’s account with the Consultant, for any payment received by the Consultant after 30 calendar days from the date of the invoice, excepting any portion of the invoiced amount in dispute and resolved in favor of the Client. The adjusted previous balance is the amount owed the Consultant on the preceding invoice, less payments and credits received. Application of this service charge as a consequence of late payment does not constitute any willingness of the Consultant to finance the Client’s operations, and no such willingness shall be inferred. The Consultant reserves the right to suspend services without notice on any and all of Clients projects under contract, when any amount is over 30 days past due. ",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "PART 4 - SUPPLEMENTAL CONDITIONS",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "4.1.	STATE SPECIFIC CHARGES:  LYNXSPRING charges “User Tax” for services rendered in the State of Missouri.  More Municipalities will be adding their version as time progresses.  This will result in additional fee’s to cover cost for them as well and LYNXSPRING will notify you of associated cost as the requirements are discovered.  LYNXSPRING reserves the right to charge for all additional cost as new Municipalities requirements are encountered.  LYNXSPRING will notify the client immediately upon determination of any extra cost.",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                }),

                new Paragraph({
                    text: "4.2.	HOURLY RATE SCHEDULE: As of August 24, 2017",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                    spacing: {
                        after: 200,
                    },
                }),

                tableadmin,

                new Paragraph({
                    text: "Note: Travel per diem is based on a ten-hour day (IE:  Engineer per diem is $250.00 X 10 = $2,500.00)",
                    style: "normalPara4",
                    alignment: AlignmentType.LEFT,
                    spacing: {
                        after: 7000,
                    },
                }),


                new Paragraph({
                    text: "LYNXSPRING SCHEDULE OF INVOICES AND STATEMENT OF WORK",
                    style: "header3",
                    alignment: AlignmentType.CENTER
                }),


                new Paragraph({
                    text: "STATEMENT OF WORK FOR:  LTF – MIDDLETOWN, NJ ",
                    style: "normalPara7",
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        after: 200,
                    },
                }),

                new Paragraph({
                    text: "MILESTONES AND DELIVERABLES INCLUDE:",
                    style: "normalPara7",
                    alignment: AlignmentType.CENTER,
                    spacing: {
                        after: 400,
                    },
                }),


                lasttable,

                total,

            ],
        },],
    });



    const b64string = await Packer.toBase64String(doc);

    res.setHeader('Content-Disposition', 'attachment; filename=NWwellsFargo.docx');
    res.send(Buffer.from(b64string, 'base64'));
});


app.post('/api/excel', cors(), function (req, res) {

    var chunks = [];
    res.on("data", function (chunk) {
        chunks.push(chunk);
    });
    res.on("end", function (chunk) {
        var body = Buffer.concat(chunks);
        // console.log(body.toString());
    });
    res.on("error", function (error) {
        // console.error(error);
    });
    if (isEmpty(req.body)) {
        return;
    }

    let wellsfargo = req.body;

    let workbook = new excel.Workbook();

    let worksheet = workbook.addWorksheet('Quote Form');

    // border none
    // worksheet.views = [{ showGridLines: false }];

    worksheet.getRow(1).height = 20;
    worksheet.getRow(2).height = 20;
    worksheet.getRow(3).height = 20;
    worksheet.getRow(4).height = 20;
    worksheet.getRow(5).height = 20;
    worksheet.getRow(6).height = 20;
    worksheet.getRow(7).height = 20;
    worksheet.getRow(8).height = 20;
    worksheet.getRow(9).height = 20;
    worksheet.getRow(10).height = 20;
    worksheet.getRow(11).height = 20;
    worksheet.getRow(12).height = 20;

    worksheet.columns = [{ key: 'A', width: 8.0 }, { key: 'B', width: 10.0 }, { key: 'C', width: 15.0 },
    { key: 'D', width: 25.0 }, { key: 'E', width: 40.0 }, { key: 'E', width: 15.0 }, { key: 'F', width: 15.0 },
    { key: 'H', width: 10.0 }
    ];

    const imageId1 = workbook.addImage({
        filename: './images/lynxspring.png',
        extension: 'png',
    });

    worksheet.addImage(imageId1, 'F1:H5',);


    worksheet.mergeCells('A1:B1');
    worksheet.getCell('A1:B1').value = 'Quote Form';
    worksheet.getCell('A1:B1').font = {
        size: 14,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell('A1:B1').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };

    worksheet.getCell('A1:B1').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('A1:B1').border = {
        right: { style: 'thin' }
    };

    worksheet.mergeCells('C1');
    worksheet.getCell('C1').value = wellsfargo.quoteForm;
    worksheet.getCell('C1').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('C1').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('C1').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DDEBF7' },
        bgColor: { argb: 'DDEBF7' }
    };
    worksheet.getCell('C1').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('C1').border = {

    };

    worksheet.mergeCells('D1');
    worksheet.getCell('D1').value = '';
    worksheet.getCell('D1').font = {
        size: 14,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('D1').border = {
        left: { style: 'thin' },
    };
    worksheet.getCell('D1').alignment = { vertical: 'middle', horizontal: 'center', };

    worksheet.mergeCells('E1:E5');
    worksheet.getCell('E1:E5').value = "2900 NE Independence Avenue.\nLee's Summit, MO 64060.\nP: 816.347.3500 | F: 816.875 5642";
    worksheet.getCell('E1:E5').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        bold: true

    };
    worksheet.getCell('E1:E5').alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

    worksheet.mergeCells('A2:B2');
    worksheet.getCell('A2:B2').value = 'Lynxspring Inc.';
    worksheet.getCell('A2:B2').font = {
        size: 12,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };

    worksheet.getCell('A2:B2').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('A2:B2').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('A2:B2').border = {
        right: { style: 'thin' }
    };

    worksheet.mergeCells('C2');
    worksheet.getCell('C2').value = wellsfargo.lynxspringInc;
    worksheet.getCell('C2').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('C2').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DDEBF7' },
        bgColor: { argb: 'DDEBF7' }
    };
    worksheet.getCell('C2').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('C2').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };

    worksheet.mergeCells('D2');
    worksheet.getCell('D2').value = '';
    worksheet.getCell('D2').font = {
        size: 12,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('D2').alignment = { vertical: 'center', horizontal: 'left' };



    worksheet.mergeCells('A3:B3');
    worksheet.getCell('A3:B3').value = 'QUOTE TO:';
    worksheet.getCell('A3:B3').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };

    worksheet.getCell('A3:B3').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('A3:B3').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('A3:B3').border = {
        right: { style: 'thin' }
    };

    worksheet.mergeCells('C3');
    worksheet.getCell('C3').value = wellsfargo.quoteTo;
    worksheet.getCell('C3').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('C3').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DDEBF7' },
        bgColor: { argb: 'DDEBF7' }
    };
    worksheet.getCell('C3').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('C3').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('D3');
    worksheet.getCell('D3').value = '';
    worksheet.getCell('D3').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('D3').alignment = { vertical: 'middle', horizontal: 'left' };



    worksheet.mergeCells('A4:B4');
    worksheet.getCell('A4:B4').value = 'Name';
    worksheet.getCell('A4:B4').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };

    worksheet.getCell('A4:B4').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('A4:B4').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('A4:B4').border = {
        right: { style: 'thin' }
    };

    worksheet.mergeCells('C4');
    worksheet.getCell('C4').value = wellsfargo.name;
    worksheet.getCell('C4').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('C4').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DDEBF7' },
        bgColor: { argb: 'DDEBF7' }
    };
    worksheet.getCell('C4').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('C4').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };


    worksheet.mergeCells('D4');
    worksheet.getCell('D4').value = "";
    worksheet.getCell('D4').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('D4').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('A5:B5');
    worksheet.getCell('A5:B5').value = 'Project';
    worksheet.getCell('A5:B5').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };

    worksheet.getCell('A5:B5').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('A5:B5').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('A5:B5').border = {
        right: { style: 'thin' }
    };


    worksheet.mergeCells('C5');
    worksheet.getCell('C5').value = wellsfargo.project;
    worksheet.getCell('C5').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('C5').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DDEBF7' },
        bgColor: { argb: 'DDEBF7' }
    };
    worksheet.getCell('C5').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('C5').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('D5');
    worksheet.getCell('D5').value = "";
    worksheet.getCell('D5').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('D5').alignment = { vertical: 'middle', horizontal: 'center' };


    worksheet.mergeCells('A6:B6');
    worksheet.getCell('A6:B6').value = 'Date:';
    worksheet.getCell('A6:B6').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };

    worksheet.getCell('A6:B6').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('A6:B6').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('A6:B6').border = {
        right: { style: 'thin' }
    };


    worksheet.mergeCells('C6');
    worksheet.getCell('C6').value = wellsfargo.date;
    worksheet.getCell('C6').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('C6').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DDEBF7' },
        bgColor: { argb: 'DDEBF7' }
    };
    worksheet.getCell('C6').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('C6').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('D6');
    worksheet.getCell('D6').value = "";
    worksheet.getCell('D6').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('D6').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('E6');
    worksheet.getCell('E6').value = "";
    worksheet.getCell('E6').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('E6').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('F6');
    worksheet.getCell('F6').value = "";
    worksheet.getCell('F6').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('F6').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('G6');
    worksheet.getCell('G6').value = "";
    worksheet.getCell('G6').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('G6').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('H6');
    worksheet.getCell('H6').value = "";
    worksheet.getCell('H6').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('H6').alignment = { vertical: 'middle', horizontal: 'center' };


    worksheet.mergeCells('A7:B7');
    worksheet.getCell('A7:B7').value = 'Expires';
    worksheet.getCell('A7:B7').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };

    worksheet.getCell('A7:B7').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('A7:B7').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('A7:B7').border = {
        right: { style: 'thin' }
    };

    worksheet.mergeCells('C7');
    worksheet.getCell('C7').value = wellsfargo.expires;
    worksheet.getCell('C7').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('C7').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DDEBF7' },
        bgColor: { argb: 'DDEBF7' }
    };
    worksheet.getCell('C7').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('C7').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };



    worksheet.mergeCells('D7:E9');
    worksheet.getCell('D7:E9').value = 'All Funds in US $';
    worksheet.getCell('D7:E9').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };

    worksheet.getCell('D7:E9').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('D7:E9').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('D7:E9').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };


    worksheet.mergeCells('F7');
    worksheet.getCell('F7').value = "Multiplier";
    worksheet.getCell('F7').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('F7').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('F7').alignment = { vertical: 'middle', horizontal: 'left' };

    worksheet.mergeCells('G7:H7');
    worksheet.getCell('G7:H7').value = wellsfargo.multiplier;
    worksheet.getCell('G7:H7').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('G7:H7').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('G7:H7').alignment = { vertical: 'left', horizontal: 'center' };
    worksheet.getCell('G7:H7').border = {};


    worksheet.mergeCells('A8:B8');
    worksheet.getCell('A8:B8').value = 'Payment terms';
    worksheet.getCell('A8:B8').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };

    worksheet.getCell('A8:B8').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('A8:B8').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('A8:B8').border = {
        right: { style: 'thin' }
    };

    worksheet.mergeCells('C8');
    worksheet.getCell('C8').value = wellsfargo.paymentTerms;
    worksheet.getCell('C8').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('C8').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DDEBF7' },
        bgColor: { argb: 'DDEBF7' }
    };

    worksheet.getCell('C8').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('C8').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('F8');
    worksheet.getCell('F8').value = "";
    worksheet.getCell('F8').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('F8').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('F8').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('G8');
    worksheet.getCell('G8').value = "";
    worksheet.getCell('G8').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('G8').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('G8').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('H8');
    worksheet.getCell('H8').value = "";
    worksheet.getCell('H8').font = {
        size: 11,
        name: 'Calibri',
        family: 1

    };
    worksheet.getCell('H8').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('H8').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('A9:B9');
    worksheet.getCell('A9:B9').value = 'Prepared by';
    worksheet.getCell('A9:B9').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };

    worksheet.getCell('A9:B9').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('A9:B9').alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    worksheet.getCell('A9:B9').border = {
        right: { style: 'thin' },
        bottom: { style: 'thin' }
    };

    worksheet.mergeCells('C9');
    worksheet.getCell('C9').value = wellsfargo.preparedBy;
    worksheet.getCell('C9').font = {
        size: 11,
        name: 'Calibri',
        family: 1,

    };
    worksheet.getCell('C9').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DDEBF7' },
        bgColor: { argb: 'DDEBF7' }
    };
    worksheet.getCell('C9').alignment = { vertical: 'bottom', horizontal: 'left', wrapText: true };
    worksheet.getCell('C9').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };

    worksheet.mergeCells('F9:H9');
    worksheet.getCell('F9:H9').value = 'ALL PRICES IN US$';
    worksheet.getCell('F9:H9').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('F9:H9').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('F9:H9').alignment = { vertical: 'middle', horizontal: 'left' };
    worksheet.getCell('F9:H9').border = {};

    worksheet.mergeCells('A10:E10');
    worksheet.getCell('A10:E10').value = 'BILL OF MATERIAL';
    worksheet.getCell('A10:E10').font = {
        size: 6.5,
        name: 'Arial Narrow',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('A10:E10').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('A10:E10').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('A10:E10').alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getCell('A10:E10').border = {};


    worksheet.mergeCells('F10');
    worksheet.getCell('F10').value = "List Price";
    worksheet.getCell('F10').font = {
        size: 10,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('F10').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('F10').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('F10').alignment = { vertical: 'middle', horizontal: 'left' };

    worksheet.mergeCells('G10');
    worksheet.getCell('G10').value = "Net Price";
    worksheet.getCell('G10').font = {
        size: 10,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('G10').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('G10').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('G10').alignment = { vertical: 'middle', horizontal: 'left' };

    worksheet.mergeCells('H10');
    worksheet.getCell('H10').value = "Extended";
    worksheet.getCell('H10').font = {
        size: 10,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('H10').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('H10').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('H10').alignment = { vertical: 'middle', horizontal: 'left' };


    worksheet.mergeCells('A11');
    worksheet.getCell('A11').value = "ITEM";
    worksheet.getCell('A11').font = {
        size: 6.5,
        name: 'Arial Narrow',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('A11').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('A11').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('A11').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('B11');
    worksheet.getCell('B11').value = "QTY";
    worksheet.getCell('B11').font = {
        size: 6.5,
        name: 'Arial Narrow',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('B11').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('B11').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('B11').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('C11');
    worksheet.getCell('C11').value = "VENDOR";
    worksheet.getCell('C11').font = {
        size: 6.5,
        name: 'Arial Narrow',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('C11').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('C11').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('C11').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('D11');
    worksheet.getCell('D11').value = "PART NO";
    worksheet.getCell('D11').font = {
        size: 6.5,
        name: 'Arial Narrow',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('D11').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('D11').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('D11').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('E11');
    worksheet.getCell('E11').value = "DESCRIPTION";
    worksheet.getCell('E11').font = {
        size: 6.5,
        name: 'Arial Narrow',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('E11').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('E11').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('E11').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('F11');
    worksheet.getCell('F11').value = "";
    worksheet.getCell('F11').font = {
        size: 6.5,
        name: 'Arial Narrow',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('F11').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('F11').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('F11').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('G11');
    worksheet.getCell('G11').value = "";
    worksheet.getCell('G11').font = {
        size: 6.5,
        name: 'Arial Narrow',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('G11').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('G11').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('G11').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('H11');
    worksheet.getCell('H11').value = "";
    worksheet.getCell('H11').font = {
        size: 6.5,
        name: 'Arial Narrow',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('H11').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('H11').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('H11').alignment = { vertical: 'middle', horizontal: 'center' };


    worksheet.mergeCells('A12');
    worksheet.getCell('A12').value = "";
    worksheet.getCell('A12').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('A12').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('A12').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('A12').alignment = { vertical: 'middle', horizontal: 'center' };


    worksheet.mergeCells('B12');
    worksheet.getCell('B12').value = "";
    worksheet.getCell('B12').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('B12').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('B12').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('B12').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('C12');
    worksheet.getCell('C12').value = "";
    worksheet.getCell('C12').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('C12').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('C12').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('C12').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('D12');
    worksheet.getCell('D12').value = "";
    worksheet.getCell('D12').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('D12').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('D12').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('D12').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('E12');
    worksheet.getCell('E12').value = "";
    worksheet.getCell('E12').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('E12').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('E12').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('E12').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('F12');
    worksheet.getCell('F12').value = "";
    worksheet.getCell('F12').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('F12').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('F12').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('F12').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('G12');
    worksheet.getCell('G12').value = "";
    worksheet.getCell('G12').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('G12').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('G12').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('G12').alignment = { vertical: 'middle', horizontal: 'center' };

    worksheet.mergeCells('H12');
    worksheet.getCell('H12').value = "";
    worksheet.getCell('H12').font = {
        size: 11,
        name: 'Calibri',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true

    };
    worksheet.getCell('H12').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('H12').border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
    worksheet.getCell('H12').alignment = { vertical: 'middle', horizontal: 'center' };

    for (let i = 0; i < wellsfargo.installColumns.length; i++) {
        let temp = i + 13;

        worksheet.mergeCells('A' + temp);
        worksheet.getCell('A' + temp).value = wellsfargo.installColumns[i].coloumn1;
        worksheet.getCell('A' + temp).font = {
            size: 11,
            name: 'Calibri',
            family: 1,


        };

        worksheet.getCell('A' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell('A' + temp).alignment = { vertical: 'middle', horizontal: 'center' };

        worksheet.mergeCells('B' + temp);
        worksheet.getCell('B' + temp).value = wellsfargo.installColumns[i].coloumn2;
        worksheet.getCell('B' + temp).font = {
            size: 11,
            name: 'Calibri',
            family: 1,


        };

        worksheet.getCell('B' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell('B' + temp).alignment = { vertical: 'middle', horizontal: 'center' };

        worksheet.mergeCells('C' + temp);
        worksheet.getCell('C' + temp).value = wellsfargo.installColumns[i].coloumn3;
        worksheet.getCell('C' + temp).font = {
            size: 11,
            name: 'Calibri',
            family: 1,


        };

        worksheet.getCell('C' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell('C' + temp).alignment = { vertical: 'middle', horizontal: 'center' };


        worksheet.mergeCells('D' + temp);
        worksheet.getCell('D' + temp).value = wellsfargo.installColumns[i].coloumn4;
        worksheet.getCell('D' + temp).font = {
            size: 11,
            name: 'Calibri',
            family: 1,


        };

        worksheet.getCell('D' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell('D' + temp).alignment = { vertical: 'middle', horizontal: 'left' };


        worksheet.mergeCells('E' + temp);
        worksheet.getCell('E' + temp).value = wellsfargo.installColumns[i].coloumn5;
        worksheet.getCell('E' + temp).font = {
            size: 11,
            name: 'Calibri',
            family: 1,


        };

        worksheet.getCell('E' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell('E' + temp).alignment = { vertical: 'middle', horizontal: 'left' };

        worksheet.mergeCells('F' + temp);
        worksheet.getCell('F' + temp).value = wellsfargo.installColumns[i].coloumn6;
        worksheet.getCell('F' + temp).font = {
            size: 11,
            name: 'Calibri',
            family: 1,


        };

        worksheet.getCell('F' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell('F' + temp).alignment = { vertical: 'middle', horizontal: 'right' };

        worksheet.mergeCells('G' + temp);
        worksheet.getCell('G' + temp).value = wellsfargo.installColumns[i].coloumn7;
        worksheet.getCell('G' + temp).font = {
            size: 11,
            name: 'Calibri',
            family: 1,


        };

        worksheet.getCell('G' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell('G' + temp).alignment = { vertical: 'middle', horizontal: 'right' };

        worksheet.mergeCells('H' + temp);
        worksheet.getCell('H' + temp).value = wellsfargo.installColumns[i].coloumn8;
        worksheet.getCell('H' + temp).font = {
            size: 11,
            name: 'Calibri',
            family: 1,


        };

        worksheet.getCell('H' + temp).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
        worksheet.getCell('H' + temp).alignment = { vertical: 'middle', horizontal: 'right' };

    }

    let installColumnslength = wellsfargo.installColumns.length + 13;
    let mergeCellAlias = 'A' + installColumnslength + ':' + 'F' + installColumnslength;

    worksheet.mergeCells(mergeCellAlias);
    worksheet.getCell(mergeCellAlias).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };

    worksheet.mergeCells('G' + installColumnslength);
    worksheet.getCell('G' + installColumnslength).value = 'Total';
    worksheet.getCell('G' + installColumnslength).font = {
        size: 9,
        name: 'Arial',
        family: 1,
        color: { argb: 'FFFFFF' },
        bold: true
    };
    worksheet.getCell('G' + installColumnslength).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('G' + installColumnslength).alignment = { vertical: 'middle', horizontal: 'right' };

    worksheet.mergeCells('H' + installColumnslength);
    worksheet.getCell('H' + installColumnslength).value = wellsfargo.total;
    worksheet.getCell('H' + installColumnslength).font = {
        size: 9,
        name: 'Arial',
        family: 1,
        bold: true,
        color: { argb: 'FFFFFF' },
    };
    worksheet.getCell('H' + installColumnslength).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '0070C0' },
        bgColor: { argb: '0070C0' }
    };
    worksheet.getCell('H' + installColumnslength).alignment = { vertical: 'middle', horizontal: 'right' };

    res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
        "Content-Disposition",
        "attachment; filename=" + "wellfargo" + ".xlsx"
    );
    return workbook.xlsx.write(res).then(function () {
        res['status'](200).end();
    });

});





app.listen(PORT, function () {
    console.log('App listening on port ' + PORT);
});