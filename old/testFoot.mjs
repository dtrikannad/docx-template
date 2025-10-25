// Footnotes

import * as fs from "fs";
import { Document, Header, Footer, FootnoteReferenceRun, Packer, Paragraph, TextRun, ImageRun, AlignmentType, TextWrappingType, TextWrappingSide, convertInchesToTwip, UnderlineType, Table, TableRow, TableCell, WidthType,BorderStyle, SectionType, ExternalHyperlink } from "docx";
import sizeOf from 'image-size';
import { LoremIpsum } from "lorem-ipsum";

function calcSizeFromInches(inches) {
    return Math.round(inches * (275/1.16));
}

let fn = {};
fn['1'] = { children: [new Paragraph({style:"Footnote", children: [
    new TextRun(" "),
    new TextRun("To note, Ms. Allen’s husband was the driver of the vehicle, and he is not presenting a bodily injury claim, at least to the best of my knowledge.")]})] };
fn['2'] = { children: [new Paragraph({style:"Footnote", children: [
    new TextRun(" "),
    new TextRun("Ms. Allen continued to receive massages when her pain flared up as indicated in the enclosed receipts, and she will continue to do so for an indefinite period of time.")]})] };
fn['3'] = { children: [new Paragraph("Testing 1,2,3")] };

let t1 = "This letter shall serve as a statement of Kristine Allen's damages as a result of the above-referenced loss. Enclosed for your review, please find a copy of the medical and billing records reflecting the treatment she received as a result of her collision caused by your insured"

const imagePath = './data/Picture1.jpg';
const imageBuffer = fs.readFileSync(imagePath);

const dimensions = sizeOf(imageBuffer);
const originalWidth = dimensions.width;
const originalHeight = dimensions.height;
const desiredWidth = calcSizeFromInches(1.16);
const proportionalHeight = (desiredWidth / originalWidth) * originalHeight;

const dateOptions = {
  month: 'long', // Full month name (e.g., "January")
  day: 'numeric', // Day of the month (e.g., "12")
  year: 'numeric' // Full year (e.g., "2025")
};

const defaultProperties = {
    page: {
        margin: {
            top: convertInchesToTwip(.25),
            bottom: convertInchesToTwip(.5),
            header: convertInchesToTwip(0.2), 
            footer: convertInchesToTwip(0.4),
        },
        size: {
            height: convertInchesToTwip(11),
            width: convertInchesToTwip(8.5)
        },
        // footer: 1000,
        // header: 5000
    },
    titlePage: true,
};

const defaultFooter = {
    // default: new Footer({
    //     children: [
    //         new Paragraph({}),
    //         new Paragraph({
    //             style: "Footer",
    //             text: "650 Town Center Drive, Suite 1400 ǁ Costa Mesa, California 92626"
    //         }),
    //         new Paragraph({
    //             style: "Footer",
    //             text: "Phone: 714.955.4551 ǁ Fax: 714.966.0663"
    //         }),
    //         new Paragraph({
    //             style: "Footer",
    //             text: "samir@sheth-law.com"
    //         }),
    //         new Paragraph({
    //             style: "Footer",
    //             text: "www.sheth-law.com"
    //         })
    //     ]
    // }),
    first: new Footer({
        children: [
            new Paragraph({
                style: "Footer",
                text: "650 Town Center Drive, Suite 1400 ǁ Costa Mesa, California 92626"
            }),
            new Paragraph({
                style: "Footer",
                text: "Phone: 714.955.4551 ǁ Fax: 714.966.0663"
            }),
            new Paragraph({
                style: "Footer",
                children: [
                new ExternalHyperlink({
                    children: [
                        new TextRun({
                            text: "samir@sheth-law.com",
                            style: "Hyperlink",
                        }),
                    ],
                    link: "mailto:samir@sheth-law.com",
                }),
                ]
            }),
            new Paragraph({
                style: "Footer",
                children: [
                new ExternalHyperlink({
                    children: [
                        new TextRun({
                            text: "www.sheth-law.com",
                            style: "Hyperlink",
                        }),
                    ],
                    link: "https://www.sheth-law.com",
                }),
                ]
            }),
        ]
    })
};

const defaultHeader = {
                default: new Header({
                    children: [
                        new Paragraph(new Intl.DateTimeFormat('en-US', dateOptions).format(new Date())),
                        new Paragraph({})
                    ],
                })
            };

const reTable = new Table({
    indent: {
        size: 600,
        type: WidthType.DXA,
    },
    borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.NONE },
        left: { style: BorderStyle.NONE},
        right: { style: BorderStyle.NONE}
    },
    rows: [
        new TableRow({
            children: [
                new TableCell({
                                    borders: {
                                        top: {
                                            style: BorderStyle.NONE,
                                            size: 3,
                                            color: "FF0000",
                                        },
                                        bottom: {
                                            style: BorderStyle.NONE,
                                            size: 3,
                                            color: "0000FF",
                                        },
                                        left: {
                                            style: BorderStyle.NONE,
                                            size: 3,
                                            color: "00FF00",
                                        },
                                        right: {
                                            style: BorderStyle.NONE,
                                            size: 3,
                                            color: "#ff8000",
                                        },
                                    },
                    children: [new Paragraph({
                        children: [
                            new TextRun({
                                text:"Re:",
                                    bold: true,
                                    italics: true
                            })
                        ],
                        alignment: AlignmentType.RIGHT,
                    })],
                    width: {
                        size: convertInchesToTwip(.55),
                        type: WidthType.DXA
                    },
                }),
                new TableCell({
                                    borders: {
                                        top: { style: BorderStyle.NONE,
                                            size: 3,
                                            color: "0000FF"
                                        },                                        
                                        bottom: {
                                            style: BorderStyle.NONE,
                                            size: 3,
                                            color: "0000FF",
                                        },
                                        left: {
                                            style: BorderStyle.NONE,
                                            size: 3,
                                            color: "00FF00",
                                        },
                                        right: {
                                            style: BorderStyle.NONE,
                                            size: 3,
                                            color: "#ff8000",
                                        },
                                    },
                    width: {
                        size: convertInchesToTwip(.5),
                        type: WidthType.DXA
                    },
                    children: [new Paragraph(" ")]
                }),
                new TableCell({
                                    borders: {
                                        top: { style: BorderStyle.NONE },
                                        
                                        bottom: {
                                            style: BorderStyle.NONE,
                                            size: 3,
                                            color: "0000FF",
                                        },
                                        left: {
                                            style: BorderStyle.NONE,
                                            size: 3,
                                            color: "00FF00",
                                        },
                                        right: {
                                            style: BorderStyle.NONE,
                                            size: 3,
                                            color: "#ff8000",
                                        },
                                    },
                    children: [new Paragraph({
                        children: [
                            new TextRun({
                                text: "Kristine Allen’s Automobile Accident Dated May 24, 2024 (Claim Number: 75-68H9-93X)",
                                italics: true,
                                bold: true,
                                underline: true
                            }),
                            new TextRun({
                                text: " – Settlement Demand",
                                italics: true,
                                bold: true,
                            })
                        ]
                    })]
                }),
            ]
        })
    ]
})

console.log(`Original Width: ${originalWidth}\nOriginal Height: ${originalHeight}\nDesired Width: ${desiredWidth}`)


const headerLogo = new ImageRun({
    type: 'jpg',
    data: imageBuffer,
    transformation: {
        width: desiredWidth,
        height: proportionalHeight
    },
    // floating: {
    //     // horizontalPosition: {
    //     //     offset: 2014400,
    //     // },
    //     // verticalPosition: {
    //     //     offset: 2014400,
    //     // },
    //     // wrap: {
    //     //     type: TextWrappingType.SQUARE,
    //     //     side: TextWrappingSide.BOTH_SIDES,
    //     // },
    //     // margins: {
    //     //     top: 201440,
    //     //     bottom: 201440,
    //     // },
    // },
});

const doc = new Document({
    footnotes: fn,    
    styles: {
        default: {
            document: {
                run: {
                    size: "12pt",
                    font: "Times New Roman"
                }
            }
        },
        paragraphStyles: [
            {
                name:"CenterBoldItalics",
                basedOn: "Normal",
                next:"Normal",
                quickFormat: true,
                paragraph: {
                    // spacing: {
                    //     after: 240
                    // },
                    // indent: {
                    //     firstLine: 720
                    // },
                    alignment: AlignmentType.CENTER
                },
                run: {
                    bold: true,
                    italics: true,
                    underline: true
                }
            },
            {
                name:"LeftBoldItalics",
                basedOn: "Normal",
                next:"Normal",
                quickFormat: true,
                paragraph: {
                    // spacing: {
                    //     after: 240
                    // },
                    // indent: {
                    //     firstLine: 720
                    // },
                    alignment: AlignmentType.LEFT
                },
                run: {
                    bold: true,
                    italics: true,
                    underline: true
                }
            },
            {
                name:"Paragraph",
                basedOn: "Normal",
                next:"Normal",
                quickFormat: true,
                paragraph: {
                    spacing: {
                        after: 300
                    },
                    indent: {
                        firstLine: 720
                    },
                    alignment: AlignmentType.JUSTIFIED
                }
            },
            {
                name: "Footnote",
                paragraph: {
                    spacing: {
                        after: 240
                    },
                    alignment: AlignmentType.LEFT
                },
                run: {
                    font: "Calibri",
                    size: "10pt",
                }
            },
            {
                name: "Footer",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                paragraph: {
                    alignment: AlignmentType.CENTER
                },
                run: {
                    size: "10pt"
                }
            }
        ]
    },
    sections: [
        {
            properties: defaultProperties,
            footers: defaultFooter,
            headers: defaultHeader,
            children: [
                new Paragraph({
                    children: [headerLogo],
                    alignment: AlignmentType.CENTER
                }),
                new Paragraph({}),
                new Paragraph({
                    children: [
                        new TextRun(new Intl.DateTimeFormat('en-US', dateOptions).format(new Date()))
                    ],
                    alignment: AlignmentType.CENTER
                }),
                new Paragraph({}),
                new Paragraph({
                    text:"PURSUANT TO EVIDENCE CODE §§ 1152 AND 1154",
                    style: "CenterBoldItalics"
                }),
                new Paragraph({}),
                new Paragraph({}),
                new Paragraph({
                    text:"VIA ELECTRONIC MAIL ONLY (statefarmclaims@statefarm.com)",
                    style: "LeftBoldItalics"
                }),
                new Paragraph({}),
                new Paragraph("David Mulholland"),
                new Paragraph("State Farm"),
                new Paragraph("P.O. Box 106171"),
                new Paragraph("Atlanta, Georgia 30348"),
                new Paragraph({}),
                reTable,
                new Paragraph({}),
                new Paragraph("Dear Mr. Mulholland:"),
                new Paragraph({}),
                new Paragraph({
                    text: "This letter shall serve as a statement of Kristine Allen’s damages as a result of the above-referenced loss.  Enclosed for your review, please find a copy of medical and billing records reflecting the treatment she received as a result of her collision caused by your insured.",
                    style: "Paragraph"
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                    text: "On May 24, 2024, Ms. Allen was the front passenger in a vehicle driving on Alton Parkway in the city of Irvine, state of California.",
                        }),
                        new FootnoteReferenceRun(1),
                        new TextRun("  "),
                        new TextRun({
                            text: "As the vehicle approached the intersection of Alton Parkway and Irvine Boulevard, it stopped for the red light that it was faced it, and, shortly thereafter, it was forcefully struck in the rear by your insured’s vehicle. As a result of the force and nature of the impact, Ms. Allen sustained bodily injuries necessitating medical attention."
                        })
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "After the accident occurred, Ms. Allen began to experience pain and soreness throughout her body which grew progressively worse, so she went to a massage therapist the following day to see if a massage might provide her some relief."
                        }),
                        new FootnoteReferenceRun(2),
                        new TextRun('  '),
                        new TextRun("After realizing that the massage only provided temporary relief and that her pain continued to linger and heighten on occasion, she visited Dr. John Chen at Compassion Chiropractic for an evaluation of his injuries. At the time of her initial evaluation, Ms. Allen complained of dull and tight pain in her neck, dull pain and tightness in her shoulders, and dull and achy pain in her back, all of which she rated at a moderate level.")
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    children: [
                        new TextRun("Upon conducting his evaluation, Dr. Chen determined that in addition to being slow to respond, being in acute distress, and having guarded movements, Ms. Allen was suffering from tenderness, muscle spasms, and a limited range of motion in her cervical region; tenderness and muscle spasms in her thoracic region; and tenderness, muscle spasms, and a severely limited range of motion in her lumbar region. Further, Ms. Allen tested positively when the following orthopedic tests were performed: Cervical Compression, Distraction, Shoulder Depression, Single Leg Raise, Bilateral Leg Raise, Milgram’s, and Kemp’s. Finally, x-rays taken of Ms. Allen’s cervical and lumbar spine revealed loss of her cervical lordosis. Ms. Allen was diagnosed with the following:")
                    ],
                    style: "Paragraph"
                })
            ]
        },
        // {
        //     children: [
        //         new Paragraph({
        //             children: [
        //                 new TextRun({
        //                     children: ["Hello"],
        //                 }),
        //                 new FootnoteReferenceRun(1),
        //                 new TextRun({
        //                     children: [" World!"],
        //                 }),
        //                 new FootnoteReferenceRun(2),
        //             ],
        //         }),
        //     ],
        // },
    ],
});



Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});