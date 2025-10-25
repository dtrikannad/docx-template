// Footnotes

import * as fs from "fs";
import { 
    Document, 
    Header, 
    Footer, 
    FootnoteReferenceRun, 
    Packer, 
    Paragraph, 
    TextRun, 
    ImageRun, 
    AlignmentType, 
    TextWrappingType, 
    TextWrappingSide, 
    convertInchesToTwip, 
    Table, 
    TableRow, 
    TableCell, 
    WidthType,BorderStyle, 
    ExternalHyperlink, 
    PageNumber, 
    HorizontalPositionRelativeFrom,
    VerticalPositionRelativeFrom,
    LevelFormat, 
    PageBreak 
} from "docx";
import sizeOf from 'image-size';
import imageType, {minimumBytes} from 'image-type';

function calcSizeFromInches(inches) {
    return Math.round(inches * (275/1.16));
}

let fn = {};
fn['1'] = { 
    children: [
        new Paragraph(
            {
                style:"Footnote", 
                children: [
                    new TextRun(" "),
                    new TextRun("To note, Ms. Allen’s husband was the driver of the vehicle, and he is not presenting a bodily injury claim, at least to the best of my knowledge.")
                ]
            }
        )
    ]
};
fn['2'] = { children: [new Paragraph({style:"Footnote", children: [
    new TextRun(" "),
    new TextRun("Ms. Allen continued to receive massages when her pain flared up as indicated in the enclosed receipts, and she will continue to do so for an indefinite period of time.")]})] };
fn['3'] = { children: [new Paragraph({style:"Footnote", children: [
    new TextRun(" "),
    new TextRun("The charges for the late rescheduling noted on the invoice have been removed from the figure noted above.")]})] };

const imagePath = './data/Picture1.jpg';
const imageBuffer = fs.readFileSync(imagePath);
const logoImageType = await imageType(imageBuffer);

const dimensions = sizeOf(imageBuffer);
const originalWidth = dimensions.width;
const originalHeight = dimensions.height;
const desiredWidth = calcSizeFromInches(1.16);
const proportionalHeight = (desiredWidth / originalWidth) * originalHeight;
const desiredHeaderHeight = calcSizeFromInches(0.25);
const proportionalHeaderWidth = Math.round((desiredHeaderHeight / originalHeight) * originalWidth);

const dateOptions = {
  month: 'long', // Full month name (e.g., "January")
  day: 'numeric', // Day of the month (e.g., "12")
  year: 'numeric' // Full year (e.g., "2025")
};

const firstPageLogo = new ImageRun({
    type: logoImageType.ext,
    data: imageBuffer,
    transformation: {
        width: desiredWidth,
        height: proportionalHeight
    },
});

const headerLogo = new ImageRun({
    type: 'jpg',
    data: imageBuffer,
    transformation: {
        width: proportionalHeaderWidth,
        height: desiredHeaderHeight
    },
    floating: {
        horizontalPosition: {
            relative: HorizontalPositionRelativeFrom.RIGHT_MARGIN,
            offset: -((proportionalHeaderWidth)*10000),
        },
        verticalPosition: {
            relative: VerticalPositionRelativeFrom.TOP_MARGIN,
            offset: 230000,
            // offset: convertInchesToTwip(1.5)*100
        },
        wrap: {
            type: TextWrappingType.SQUARE,
            side: TextWrappingSide.BOTH_SIDES,
        },
        margins: {
            top: 0,
            bottom: 0,
            left: 0,
            right: 0
        },
        behindDocument: true
    },
});

const defaultProperties = {
    page: {
        margin: {
            top: convertInchesToTwip(.25),
            bottom: convertInchesToTwip(.5),
            header: convertInchesToTwip(0.4), 
            footer: convertInchesToTwip(0.4),
        },
        size: {
            height: convertInchesToTwip(11),
            width: convertInchesToTwip(8.5)
        },
    },
    titlePage: true,
};

const defaultFooter = {
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
                new TextRun("samir@sheth-law.com"),
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
                            bold: true
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
            new Paragraph({
                children: [
                    new TextRun({
                        children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES],
                    })
                ]
            }),
            new Paragraph({
                children: [headerLogo],
            }),
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


//     const numbering = new Numbering();

// const abstractNum = numbering.createAbstractNumbering();
// abstractNum.createLevel(0, "upperRoman", "%1", "start")
//     .addParagraphProperty(new Indent(720, 260));
// abstractNum.createLevel(1, "decimal", "%2.", "start")
//     .addParagraphProperty(new Indent(1440, 980));
// abstractNum.createLevel(2, "lowerLetter", "%3)", "start")
//     .addParagraphProperty(new Indent(2160, 1700));
// const concrete = numbering.createConcreteNumbering(numberedAbstract);




const doc = new Document({
    footnotes: fn,    
    numbering: {
        config: [
            {
                reference: "my-numbering",
                levels: [
                    {
                        level: 0,
                        format: LevelFormat.DECIMAL,
                        text: "%1.",
                        alignment: AlignmentType.START,
                        style: {
                            paragraph: {
                                indent: { 
                                    left: convertInchesToTwip(1), 
                                    hanging: convertInchesToTwip(.5) 
                                },
                                spacing: {
                                    after: 300
                                },
                            },
                        },
                    },
                ],
            },
        ],
    },
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
                name: "Encl",
                paragraph: {
                    spacing: {
                        after: 240
                    },
                    alignment: AlignmentType.LEFT
                },
                run: {
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
                    children: [firstPageLogo],
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
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "A Cervical Sprain due to a Motor Vehicle Accident;",
                        })
                    ],
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 0, // Unique for each list
                    },
                }),
                new Paragraph({
                    text: "A Thoracic Spine Ligamentous Sprain due to a Motor Vehicle Accident;",
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                    },
                }),
                new Paragraph({
                    text: "Muscle Spasms in her Back;",
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                    },
                }),
                new Paragraph({
                    text: "Cervical Segmental Dysfunction;",
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                    },
                }),
                new Paragraph({
                    text: "Thoracic Segmental Dysfunction; and",
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                    },
                }),
                new Paragraph({
                    text: "Lumbar Segmental Dysfunction.",
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun("As a result of the above, Dr. Chen recommended that Ms. Allen receive care including a chiropractic rehabilitation program consisting of manual joint manipulation and/or mobilization soft tissue manipulation, manual cervical spine distraction, and therapeutic exercises.  Additionally, she was provided physical therapy modalities including electrical muscle stimulation, mechanical vibrational massage, ultrasound, mechanical traction, and mechanical inter-segmental traction. Finally, Ms. Allen was provided a home exercise program. That treatment lasted until October 4, 2024, at which time Ms. Allen reported only a slight improvement in her pain levels. She also had tightness and tenderness in the affected areas of her body.")
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    children: [
                        new TextRun("While receiving care from Dr. Chen, Ms. Allen was referred to SimonMed Imaging to have an MRI of her cervical spine performed. That MRI took place on September 5, 2024 and it revealed an ongoing loss of her cervical lordosis, mild left facet arthropathy and neural foraminal narrowing in her C3-C4 region; a 2 mm broad-based disc bulge with mild spinal canal narrowing in her C4-C5 region; a 3 mm broad-based disc extrusion in her C5-C6 region with mild to moderate spinal canal narrowing, mild ventral cord flattening, facet arthropathy, uncovertebral hypertrophy, and mild to moderate left and mild right neural foraminal narrowing; and a 1-2 mm disc bulge in her C6-C7 region.")
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    children: [
                        new TextRun("In light of the results of her MRI and her ongoing pain despite the treatment that she had been receiving, it was recommended that Ms. Allen consult with a pain management specialist.  That initial consultation took place on October 3, 2024 with Dr. Khyber Zaffarkhan at the Regenerative Institute of Newport Beach (“Regenerative”). At the time, Ms. Allen stated that her neck pain was her primary concern, he described having episodes of “nerve pain,” he said the pain was sharp, it was worse with in the mornings, it was worse on her left side than her right, it interfered with her ability to lift heavier weights, and it interfered with her ability to work normally.")
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    children: [
                        new TextRun("Her evaluation by Dr. Zaffarkhan revealed an ongoing limitation of the range of motion in her cervical spine, and she tested positively when the Facet Loading Test was administered.  After reviewing the results of her MRI with her, Dr. Zaffarkhan diagnosed Ms. Allen with the following:")
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    text: "Cervical Facet Syndrome;",
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 1, // Unique for each list
                    },
                }),
                new Paragraph({
                    text: "Cervical Herniated Discs; and",
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                    },
                }),
                new Paragraph({
                    text: "Cervical Canal Stenosis.",
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun("In light of his findings, Dr. Zaffarkhan suggested that Ms. Allen continue her care with Dr. Chen, try Extracorporeal Shockwave Therapy on 6-8 occasions, and consider receiving a Cervical Facet PRP Injection if the Shockwave Therapy proved to be ineffective.")
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    children: [
                        new TextRun("After thinking the treatment recommendations through over the course of the following week, Ms. Allen decided to move forward with the recommended Shockwave Therapy, which she received on October 29, 2024, November 5, 2024, and November 7, 2024.")
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    children: [
                        new TextRun("Ms. Allen followed up with Dr. Zaffarkhan on November 12, 2024 and stated that the three sessions mentioned above proved to be effective, so she decided to discontinue her treatment for the time being and return down the road should her neck pain flare up.")
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    children: [
                        new TextRun("As a result of Ms. Allen’s collision caused by your insured, her life and lifestyle have been negatively impacted in the following ways:")
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "She takes Crossfit classes in order to stay healthy and fit, and she had to put those classes on pause for some time while her body healed.",
                            bold: true
                        })
                    ],
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 2, // Unique for each list
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "For the few months following the accident, she constantly looked in her rearview mirror thinking that she would get struck again. That anxiety made her feel as though she could not relax at all as it spilled over into her relaxation and sleep time.",
                            bold: true
                        })
                    ],
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 2, // Unique for each list
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Her difficulty driving was exacerbated by the fact that her pain and discomfort would increase during longer commutes.",
                            bold: true
                        })
                    ],
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 2, // Unique for each list
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Her need to attend treatment sessions, and her pain otherwise, caused her to have to miss spending time with her family and attend certain functions at her daughter's school.",
                            bold: true
                        })
                    ],
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 2, // Unique for each list
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "The constant pain that she was in led her to avoid socializing with others; instead, she preferred to be alone so she could rest and recover.",
                            bold: true
                        })
                    ],
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 2, // Unique for each list
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "The above led to issues in her marriage, which resulted in her having to see a therapist to help her deal with the tension between her and her husband.",
                            bold: true
                        })
                    ],
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 2, // Unique for each list
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Her sleep was constantly interrupted, making her tired throughout the day and making her workdays that much more difficult.",
                            bold: true
                        })
                    ],
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 2, // Unique for each list
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "She had difficulty getting out of bed and dressing and grooming herself in the mornings after the accident.",
                            bold: true
                        })
                    ],
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 2, // Unique for each list
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Performing household chores and day-to-day errands (such as doing her laundry, cooking, cleaning, etc.) became difficult for her to perform, causing him to rely on her family for assistance, which was an added source of the marital tension mentioned above.",
                            bold: true
                        })
                    ],
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 2, // Unique for each list
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun("The total cost of Ms. Allen’s treatment as a result of the collision at issue is "),
                        new TextRun({
                            text: "$13,960.20",
                            bold: true
                        }),
                        new TextRun(", and is broken down as follows:")
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    text: "Massages – $545.00;",
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 3, // Unique for each list
                    },
                }),
                new Paragraph({
                    text: "Compassion Chiropractic – $4,630.00;",
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 3, // Unique for each list
                    },
                }),
                new Paragraph({
                    text: "SimonMed – $2,243.00; and",
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 3, // Unique for each list
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun("Regenerative - $6,542.20."),
                        new FootnoteReferenceRun(3),
                    ],
                    numbering: {
                        reference: "my-numbering",
                        level: 0,
                        instance: 3, // Unique for each list
                    },
                }),
                new Paragraph({
                    children: [
                        new TextRun("Factoring in the discomfort, uneasiness, and inconvenience Ms. Allen experienced, along with her pain and suffering, the nature of the trauma he sustained, the effect of the collision on her day-to-day life (as described in detail above), her likely need for future care (particularly in light of the results of her MRI), and her greater susceptibility to future injuries of the type she sustained, Ms. Allen requests "),
                        new TextRun({
                            text: "$50,000.00",
                            bold: true
                        }),
                        new TextRun(" in exchange for a full and complete release of any and all liability resulting from the above-referenced loss.")
                    ],
                    style: "Paragraph"
                }),               
                new Paragraph({
                    children: [
                        new PageBreak(), // Inserts a page break
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun("Please respond to the above request by no later than "),
                        new TextRun({
                            text: "Friday, January 3, 2025.",
                            bold: true,
                            underline: true
                        })
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    children: [
                        new TextRun("Nothing in this letter is intended to be, nor should be construed as, an admission against the interests of Kristine Allen and shall not be construed as a waiver of her rights, remedies, claims, or defenses, all of which are expressly reserved.")
                    ],
                    style: "Paragraph"
                }),  
                new Paragraph({
                    children: [
                        new TextRun("Should you have any questions or concerns, please feel free to contact me.")
                    ],
                    style: "Paragraph"
                }),  
                new Paragraph({
                    children: [
                        new TextRun("\t\t\t\t\tSincerely,")
                    ],
                    style: "Paragraph"
                }),
                new Paragraph({
                    children: [
                        new TextRun({text:"\t\t\t\t\t/s/Samir I. Sheth", bold: true, italics: true})
                    ],
                    style: "Paragraph"
                }),  
                new Paragraph({
                    children: [
                        new TextRun({text:"\t\t\t\t\tSamir I. Sheth", allCaps: true})
                    ],
                    style: "Paragraph"
                }),  
                new Paragraph({}),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "Encl.: As stated",
                        })
                    ],
                    style: "Encl"
                })
            ]
        },
    ],
});



Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});