import {
    AlignmentType,
    BorderStyle,
    convertInchesToTwip,
    Document,
    ExternalHyperlink,
    Footer,
    FootnoteReferenceRun,
    Header,
    HorizontalPositionRelativeFrom,
    ImageRun,
    LevelFormat,
    Packer,
    PageBreak,
    PageNumber,
    Paragraph,
    SectionProperties,
    Table,
    TableCell,
    TableRow,
    TextRun,
    TextWrappingSide,
    TextWrappingType,
    VerticalPositionRelativeFrom,
    WidthType,
    type IRunOptions
} from 'docx'
import sizeOf from 'image-size';
import imageType, {minimumBytes, type ImageTypeResult} from 'image-type';
import type {
    AccidentData,
    ImageDimensions,
    Footnote,
    ImageType,
} from './types';
import * as fs from 'fs';
import type { ISizeCalculationResult } from 'image-size/dist/types/interface';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
import { Command } from 'commander';

const program = new Command();
    program
        .name('write-settlement-demand')
        .description('A simple CLI that takes in a JSON data file and writes a settlement demand letter using the data.')
        .version('1.0.0')
        .option('-d, --data <path to datafile>', 'Define the path to the data file')
        .option('-l, --logo <path to image file for logo>', 'Provide the path to the image file to be used for the logo')
        .option('-o, --output <path to output>','Provide the path to the output directory')
        .option('-n, --name <name of outputfile>','provide name of the output file');

    // program.option('-v, --verbose', 'enable verbose output');
    // program.option('-n, --name <name>', 'specify a name');
    // program.option('-p, --port <port>', 'set the port number', '3000'); // Default value
    // program.option('-d, --debug', 'output extra debugging information')
    // program.option('-t, --timeout <seconds>', 'specify the timeout in seconds', '60');
    // program.option('-f, --file <path>', 'specify the file to process');
    // program.option('-g, --deepak','this is my first name')

    program.parse(process.argv);

    const options = program.opts();

    if(options.data && options.logo) {
        console.log('both required documents are provided', options.data, options.logo)
    } else {
        console.log('missing one of the required documents: data.json file and/or logo file.')
    }

    // if (options.verbose) {
    // console.log('Verbose mode is enabled.');
    // }

    // if (options.name) {
    // console.log(`Hello, ${options.name}!`);
    // }
    // console.log(`Port: ${options.port}`);

class Docx {
    private logoPath: string;
    private logoBuffer: Buffer;
    private logoImageType!: ImageTypeResult;
    private accidentDatapath: string;
    private accidentData!: AccidentData;
    private footnotes: Record<string, Footnote> = {};
    private originalLogoDimensions: ISizeCalculationResult;
    private desiredLogoDimensionsForTitle!: ImageDimensions;
    private desiredLogoDimensionsForHeader!: ImageDimensions;
    private widthOfTitleImage: number = 1.16;
    private heightOfHeaderImage: number = 0.25;
    private dateOptions: Intl.DateTimeFormatOptions =  {
        month: "long", // Full month name (e.g., "January")
        day: "numeric", // Day of the month (e.g., "12")
        year: "numeric", // Full year (e.g., "2025")
    };
    private headerLogo!: ImageRun;
    private titlePageLogo!: ImageRun;
    private defaultPageProperties = {
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
    }
    private documentSections: Array<any> = [];
    private sectionChildren: Array<any> = [];
    

    constructor(logoPath: string, accidentDataPath: string, outputFileName?: string) {
        this.logoPath = logoPath;
        
        this.accidentDatapath = accidentDataPath;
        this.logoBuffer = fs.readFileSync(this.logoPath);
        this.originalLogoDimensions = sizeOf(this.logoBuffer);
        this.desiredLogoDimensionsForTitle = { 
            width: this.calcSizeFromInches(this.widthOfTitleImage), 
            height: Math.round((this.calcSizeFromInches(this.widthOfTitleImage) / this.originalLogoDimensions.width) * this.originalLogoDimensions.height) 
        };
        this.desiredLogoDimensionsForHeader = {
            width: Math.round((this.calcSizeFromInches(this.heightOfHeaderImage) / this.originalLogoDimensions.height) * this.originalLogoDimensions.width),
            height: this.calcSizeFromInches(this.heightOfHeaderImage)
        }

    }

    // constructor(jsonData: AccidentData, logoBuffer: Buffer, headerImageBuffer?: Buffer) {
    //     this.logoBuffer = logoBuffer;
    //     this.accidentData = jsonData;


    //     this.originalLogoDimensions = sizeOf(this.logoBuffer);
    //     this.desiredLogoDimensionsForTitle = { 
    //         width: this.calcSizeFromInches(this.widthOfTitleImage), 
    //         height: Math.round((this.calcSizeFromInches(this.widthOfTitleImage) / this.originalLogoDimensions.width) * this.originalLogoDimensions.height) 
    //     };
    //     this.desiredLogoDimensionsForHeader = {
    //         width: Math.round((this.calcSizeFromInches(this.heightOfHeaderImage) / this.originalLogoDimensions.height) * this.originalLogoDimensions.width),
    //         height: this.calcSizeFromInches(this.heightOfHeaderImage)
    //     }
    //     if(headerImageBuffer) {
    //         this.headerImageBuffer = headerImageBuffer;
    //     } else {
    //         this.headerImageBuffer = logoBuffer;
    //     }
    //     this.originalHeaderImageDimensions = sizeOf(this.headerImageBuffer);
    //     this.desiredHeaderImageDimensionsForTitle = { 
    //         width: this.calcSizeFromInches(this.widthOfTitleImage), 
    //         height: Math.round((this.calcSizeFromInches(this.widthOfTitleImage) / this.originalHeaderImageDimensions.width) * this.originalHeaderImageDimensions.height) 
    //     };
    //     this.desiredHeaderImageDimensionsForHeader = {
    //         width: Math.round((this.calcSizeFromInches(this.heightOfHeaderImage) / this.originalHeaderImageDimensions.height) * this.originalLogoDimensions.width),
    //         height: this.calcSizeFromInches(this.heightOfHeaderImage)
    //     }
    // }

    private async initialize(): Promise<void> {
        this.logoImageType = (await imageType(this.logoBuffer))!

        this.accidentData = JSON.parse(fs.readFileSync(this.accidentDatapath, 'utf8'));
        if(Array.isArray(this.accidentData))
            this.accidentData = this.accidentData[0];
        this.titlePageLogo = await this.makeTitlePageLogo();
        this.headerLogo = await this.makeHeaderLogo();
    }

    private async makeTitlePageLogo(): Promise<ImageRun> {
        if(!this.logoImageType)
            await this.initialize();
        let retVal = new ImageRun(
            {
                type: this.logoImageType.ext as any,
                data: this.logoBuffer,
                transformation: {
                    width: this.desiredLogoDimensionsForTitle.width,
                    height: this.desiredLogoDimensionsForTitle.height,
                }
            }
        )
        return retVal;
    }

    private async makeHeaderLogo(): Promise<ImageRun> {
        let retVal = new ImageRun({
            type: this.logoImageType.ext as any,
            data: this.logoBuffer,
            transformation: {
                width: this.desiredLogoDimensionsForHeader.width,
                height: this.desiredLogoDimensionsForHeader.height,
            },
            floating: {
            horizontalPosition: {
                relative: HorizontalPositionRelativeFrom.RIGHT_MARGIN,
                offset: -(this.desiredLogoDimensionsForHeader.width * 10000),
            },
            verticalPosition: {
                relative: VerticalPositionRelativeFrom.TOP_MARGIN,
                offset: 230000,
            },
            wrap: {
                type: TextWrappingType.SQUARE,
                side: TextWrappingSide.BOTH_SIDES,
            },
            margins: {
                top: 0,
                bottom: 0,
                left: 0,
                right: 0,
            },
            behindDocument: true,
            },
        });
        return retVal;
    }

    private async makeDefaultFooter() {
        return {
            first: new Footer({
                children: [
                new Paragraph({
                    style: "Footer",
                    text: `${this.accidentData.attorney.streetAddress} ǁ ${this.accidentData.attorney.city}, ${this.accidentData.attorney.fullStateName} ${this.accidentData.attorney.zipCode}`,
                }),
                new Paragraph({
                    style: "Footer",
                    text: `Phone: ${this.accidentData.attorney.phoneNumber} ǁ Fax: ${this.accidentData.attorney.faxNumber}`,
                }),
                new Paragraph({
                    style: "Footer",
                    children: [new TextRun(this.accidentData.attorney.email)],
                }),
                new Paragraph({
                    style: "Footer",
                    children: [
                    new ExternalHyperlink({
                        children: [
                        new TextRun({
                            text: this.accidentData.attorney.website,
                            style: "Hyperlink",
                            bold: true,
                        }),
                        ],
                        link: `https://${this.accidentData.attorney.website}`,
                    }),
                    ],
                }),
                ],
            }),
        };

    }

    private async makeDefaultHeader() {
        if(this.headerLogo == undefined)
            await this.initialize();
        
        return {
            default: new Header({
            children: [
                new Paragraph(
                    new Intl.DateTimeFormat("en-US", this.dateOptions).format(new Date()), // Get Today's Date
                ),
                new Paragraph({
                children: [
                    new TextRun({
                    children: [
                        "Page ",
                        PageNumber.CURRENT,
                        " of ",
                        PageNumber.TOTAL_PAGES,
                    ],
                    }),
                ],
                }),
                new Paragraph({
                    children: [this.headerLogo],
                }),
            ],
            }),
        };
    }

    private async makeReTable() {
        console.log(this.accidentData.clientInfo.fullName)
        return new Table({
            indent: {
                size: 600,
                type: WidthType.DXA,
            },
            borders: {
                top: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.NONE },
                right: { style: BorderStyle.NONE },
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
                        children: [
                        new Paragraph({
                            children: [
                            new TextRun({
                                text: "Re:",
                                bold: true,
                                italics: true,
                            }),
                            ],
                            alignment: AlignmentType.RIGHT,
                        }),
                        ],
                        width: {
                        size: convertInchesToTwip(0.55),
                        type: WidthType.DXA,
                        },
                    }),
                    new TableCell({
                        borders: {
                        top: { style: BorderStyle.NONE, size: 3, color: "0000FF" },
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
                        size: convertInchesToTwip(0.5),
                        type: WidthType.DXA,
                        },
                        children: [new Paragraph(" ")],
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
                        children: [
                        new Paragraph({
                            children: [
                            new TextRun({
                                text: `${this.accidentData.clientInfo.fullName}’s Automobile Accident Dated ${new Intl.DateTimeFormat("en-US", this.dateOptions).format(new Date(this.accidentData.letterDetails.dateOfLoss))} (Claim Number: ${this.accidentData.adverseInsuranceAdjusterInfo.claimNumber})`,
                                italics: true,
                                bold: true,
                                underline: {
                                    type: "single"
                                },
                            }),
                            new TextRun({
                                text: " – Settlement Demand",
                                italics: true,
                                bold: true,
                            }),
                            ],
                        }),
                        ],
                    }),
                    ],
                }),
            ],
        });
    }

    private async makeParagraphWithLogoForTitlePage() {
        return new Paragraph({
            children: [this.titlePageLogo],
            alignment: AlignmentType.CENTER,
        });
    }   

    private makeBulletWithFootnote(bullet: string, footnoteDelimiter: string, footnoteText: string, bulletInstance: number) {
        let footnoteIndex = this.saveFootnote(footnoteText);
        const children: (TextRun | FootnoteReferenceRun)[] = [];
        let sentences = bullet
            .split(footnoteDelimiter)
            .map((sentences) => sentences.trim());
        if (sentences[0] !== undefined) {
            children.push(
                new TextRun({
                text: sentences[0],
                })
            );
        } else {
        // Optionally push a placeholder or skip this run
            children.push(new TextRun({ text: "" }));
        }
        children.push(
            new FootnoteReferenceRun(footnoteIndex),
        );

        if(sentences[1] !== undefined) {
            children.push(
                new TextRun("  "),
                new TextRun(sentences[1]),
            )
        }

        const retVal = new Paragraph({
            children,
            numbering: {
                reference: "my-numbering",
                level: 0,
                instance: bulletInstance
            }
        });
        return retVal;
    }

    private makeBoldBulletWithFootnote(bullet: string, footnoteDelimiter: string, footnoteText: string, bulletInstance: number) {
        let footnoteIndex = this.saveFootnote(footnoteText);
        const children: (TextRun | FootnoteReferenceRun)[] = [];
        let sentences = bullet
            .split(footnoteDelimiter)
            .map((sentences) => sentences.trim());
        if (sentences[0] !== undefined) {
            children.push(
                new TextRun({
                text: sentences[0],
                bold: true
                })
            );
        } else {
        // Optionally push a placeholder or skip this run
            children.push(new TextRun({ text: "", bold: true }));
        }
        children.push(
            new FootnoteReferenceRun(footnoteIndex),
        );

        if(sentences[1] !== undefined) {
            children.push(
                new TextRun({text:"  ", bold: true}),
                new TextRun({
                    text: sentences[1],
                    bold: true
                }),
            )
        }

        const retVal = new Paragraph({
            children,
            numbering: {
                reference: "my-numbering",
                level: 0,
                instance: bulletInstance
            },
            alignment: AlignmentType.JUSTIFIED
        });
        return retVal;
    }

    private makeParagraphWithFootnote(paragraph: string, footnoteDelimiter: string, footnoteText: string) {
        let footnoteIndex = this.saveFootnote(footnoteText);
        const children: (TextRun | FootnoteReferenceRun)[] = [];
        let sentences = paragraph
            .split(footnoteDelimiter)
            .map((sentences) => sentences.trim());

        if (sentences[0] !== undefined) {
            children.push(
                new TextRun({
                text: sentences[0],
                })
            );
        } else {
        // Optionally push a placeholder or skip this run
            children.push(new TextRun({ text: "" }));
        }
        children.push(
            new FootnoteReferenceRun(footnoteIndex),
        );

        if(sentences[1] !== undefined) {
            children.push(
                new TextRun("  "),
                new TextRun(sentences[1]),
            )
        }

        const retVal = new Paragraph({
            children,
            style: "Paragraph"
        });
        return retVal;
    }


    private calcSizeFromInches(inches: number) {
        return Math.round(inches * (275 / 1.16));
    }

    public static convertImageSizeFromInches(inches: number) {
        return Math.round(inches * 275/1.16);
    }
    public getNextFootnoteNumber(): string {
        let keyArray: Array<string> = Object.keys(this.footnotes);
        let nextKey: number = keyArray.length + 1;
        let key: string = nextKey.toString();
        return key;
    }
    public saveFootnote(text: string): number {
        let keyArray: Array<string> = Object.keys(this.footnotes);
        let nextKey: number = keyArray.length + 1;
        let key: string = nextKey.toString();
        this.footnotes[this.getNextFootnoteNumber()] = 
        {
            children: [
                new Paragraph(
                    {
                        style: "Footnote",
                        children: [
                            new TextRun(" "),
                            new TextRun(text)
                        ]
                    }
                )
            ]
        }
        return parseInt(key);
    }

    public getFootnotes() {
        console.log(JSON.stringify(this.footnotes, null, 2));
    }

    public async buildSectionChildren() {
        // Add Date of Document
        this.sectionChildren.push(
            await this.makeParagraphWithLogoForTitlePage(),
            new Paragraph({}),
            new Paragraph(
                {
                    children: [
                        new TextRun(new Intl.DateTimeFormat('en-US', this.dateOptions).format(new Date()))
                    ],
                    alignment: AlignmentType.CENTER
                }
            ),
            new Paragraph({}),
            new Paragraph({}),
            new Paragraph({
                text:"PURSUANT TO EVIDENCE CODE §§ 1152 AND 1154",
                style: "CenterBoldItalics"
            }),
            new Paragraph({}),
            new Paragraph({}),
            new Paragraph({
                text: `VIA ELECTRONIC MAIL ONLY (${this.accidentData.adverseInsuranceAdjusterInfo.email})`,
                style: "LeftBoldItalics"
            }),
            new Paragraph({}),
            new Paragraph(this.accidentData.adverseInsuranceAdjusterInfo.fullName),
            new Paragraph(this.accidentData.adverseInsuranceAdjusterInfo.companyName),
            new Paragraph(this.accidentData.adverseInsuranceAdjusterInfo.streetAddress),
            new Paragraph(`${this.accidentData.adverseInsuranceAdjusterInfo.city}, ${this.accidentData.adverseInsuranceAdjusterInfo.fullStateName} ${this.accidentData.adverseInsuranceAdjusterInfo.zipCode}`),
            new Paragraph({}),
            await this.makeReTable(),
            new Paragraph({}),
            new Paragraph(`Dear ${this.accidentData.adverseInsuranceAdjusterInfo.title} ${this.accidentData.adverseInsuranceAdjusterInfo.lastName}:`),
            new Paragraph({}),
        );

        this.accidentData.paragraphs.forEach(async (item, index) => {
            let child: any = {};
            switch(item.type) {
                case "encl":
                    this.sectionChildren.push(new Paragraph({
                                        children: [
                                            new TextRun({
                                                text: item.text,
                                            })
                                        ],
                                        style: "Encl"
                                    }))
                    break;
                case "pageBreak":
                    this.sectionChildren.push(new Paragraph({
                                        children: [
                                            new PageBreak(), // Inserts a page break
                                        ],
                                    }))
                    break;
                case "bulletBold":
                    if(item.text.indexOf('{footnote}') >= 0 && item.footnote != (undefined)) {
                        child = this.makeBoldBulletWithFootnote(item.text, "{footnote}",item.footnote, item.bulletInstance!);
                    } else {
                        child = new Paragraph({
                            children: [
                                new TextRun({
                                    text: item.text,
                                    bold: true
                                })
                            ],
                            numbering: {
                                reference: "my-numbering",
                                level: 0,
                                instance: item.bulletInstance!, // Unique for each list
                            },
                            alignment: AlignmentType.JUSTIFIED
                        })
                    }
                    // child.alignment = AlignmentType.JUSTIFIED;
                    this.sectionChildren.push(child);
                    break;
                case "bullet":
                    if(item.text.indexOf('{footnote}') >= 0 && item.footnote != (undefined)) {
                        child = this.makeBulletWithFootnote(item.text, "{footnote}",item.footnote, item.bulletInstance!);
                    } else {
                        child = new Paragraph({
                            text: item.text,
                            numbering: {
                                reference: "my-numbering",
                                level: 0,
                                instance: item.bulletInstance!, // Unique for each list
                            },
                        });
                    }
                    this.sectionChildren.push(child);
                    break;
                case "paragraph":
                    if(item.text.indexOf('{footnote}') >= 0 && item.footnote != (undefined)) {
                        child = this.makeParagraphWithFootnote(item.text, "{footnote}",item.footnote);
                    } else {
                        // child = new Paragraph(
                        //         {
                        //             text: item.text,
                        //             style: "Paragraph"
                        //         }
                        //     );
                        child = this.createFormattedParagraph(item.text)
                    }
                    this.sectionChildren.push(child);
                    break;
            }
        })
    }

    private createFormattedParagraph(text: string) {

        // Create a new paragraph
        const paragraph = new Paragraph({style: "Paragraph"});

        // Regular expression to match tags and text
        const regex = /(<\/?[bui]>)|([^<]+)/g;

        let isBold: boolean = false;
        let isUnderline: boolean = false;
        let isItalic: boolean = false;
        let match;

        // Process the text
        while ((match = regex.exec(text)) !== null) {
            const tag = match[1];
            const content = match[2];

            if (tag) {
                // Handle tags
                switch (tag.toLowerCase()) {
                case '<b>':
                    isBold = true;
                    break;
                case '</b>':
                    isBold = false;
                    break;
                case '<u>':
                    isUnderline = true;
                    break;
                case '</u>':
                    isUnderline = false;
                    break;
                case '<i>':
                    isItalic = true;
                    break;
                case '</i>':
                    isItalic = false;
                    break;
                }
            } else if (content) {
                // Add text with appropriate formatting
                let options: IRunOptions = {
                    text: content,
                    bold: isBold,
                    italics: isItalic,
                    ...(isUnderline && { underline: { type: "single" } })
                }
                paragraph.addChildElement(
                    new TextRun(options)
                );
            }
        }

        return paragraph;
    }
    public async buildDocument(): Promise<Document> {
        if(!this.headerLogo)
            await this.initialize();
        
        await this.buildSectionChildren();

        this.documentSections.push({
                    properties: this.defaultPageProperties,
                    footers: await this.makeDefaultFooter(),
                    headers: await this.makeDefaultHeader(),
                    children: this.sectionChildren
                });

        let retVal = new Document({
            footnotes: this.footnotes,
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
                        id:"CenterBoldItalics",
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
                            underline: {
                                type: "single"
                            },
                        },
                    },
                    {
                        id:"LeftBoldItalics",
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
                            underline: {
                                type: "single"
                            },
                        }
                    },
                    {
                        id:"Paragraph",
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
                        id:"Footnote",
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
                        id:"Encl",
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
                        id: "Footer",
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
            sections: this.documentSections
        })
        return retVal;
    }


}
function formatDateForFilename(date = new Date()) {
  return date.toISOString()
    .replace(/:/g, '-')     // Replace colons with hyphens
    .replace(/\..+$/, '')   // Remove milliseconds
    .replace('T', '_');     // Replace T with underscore
}
console.log(options);
const doc = new Docx(options.logo, options.data);
const document = await doc.buildDocument();
Packer.toBuffer(document).then((buffer) => {
    let fileName;
    if(options.name)
        fileName = options.name;
    else
        fileName = `letter_${formatDateForFilename()}.docx`;

    
        const filePath = (options.output[0] === '/') ? path.join(options.output,fileName) : path.join(__dirname,options.output,fileName);
    fs.writeFileSync(filePath, buffer);
});
// doc.getFootnotes();
