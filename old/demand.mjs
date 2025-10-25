// Demand.mjs

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
  WidthType,
  BorderStyle,
  ExternalHyperlink,
  PageNumber,
  HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom,
  LevelFormat,
  PageBreak,
} from "docx";
import sizeOf from "image-size";
import imageType from "image-type";

export class Demand {
  #name = "";
  #logoPath = "";
  #originalLogoDimensions = "";
  #logoBuffer;
  #logoImageType = {};
  #desiredLogoWidthForTitlePage = "";
  #proportionalLogoHeightForTitlePage = "";
  #desiredLogoHeightForHeader = "";
  #proportionalLogoWidthForHeader = "";
  #widthOfTitleImage = 1.16;
  #heightOfHeaderImage = 0.25;
  #footnotesObject = {};
  #sectionChildren = [];
  #dateOptions = {
    month: "long", // Full month name (e.g., "January")
    day: "numeric", // Day of the month (e.g., "12")
    year: "numeric", // Full year (e.g., "2025")
  };
  #defaultProperties = {
    page: {
      margin: {
        top: convertInchesToTwip(0.25),
        bottom: convertInchesToTwip(0.5),
        header: convertInchesToTwip(0.4),
        footer: convertInchesToTwip(0.4),
      },
      size: {
        height: convertInchesToTwip(11),
        width: convertInchesToTwip(8.5),
      },
    },
    titlePage: true,
  };

  constructor(name, logoPath) {
    this.#name = name;
    this.#logoPath = logoPath;
  }

  async #getImageMetadata() {
    this.#logoBuffer = fs.readFileSync(this.#logoPath);
    this.#originalLogoDimensions = sizeOf(this.#logoBuffer);
    this.#desiredLogoWidthForTitlePage = this.calcSizeFromInches(
      this.#widthOfTitleImage,
    );
    this.#proportionalLogoHeightForTitlePage = Math.round(
      (this.#desiredLogoWidthForTitlePage /
        this.#originalLogoDimensions.width) *
        this.#originalLogoDimensions.height,
    );
    this.#proportionalLogoWidthForHeader = Math.round(
      (this.#desiredLogoHeightForHeader / this.#originalLogoDimensions.height) *
        this.#originalLogoDimensions.width,
    );
    this.#logoImageType = await imageType(this.#logoBuffer);
  }

  async makeLetter() {
    await this.#getImageMetadata();
    let inputVariables = {
      accidentDetails: {
        dateOfLoss: "2024-05-24T00:00:00.000-08:00",
      },
      attorney: {
        email: "samir@sheth-law.com",
        website: "www.sheth-law.com",
        streetAddress: "650 Town Center Drive, Suite 1400",
        city: "Costa Mesa",
        stateCode: "CA",
        fullStateName: "California",
        zipCode: "92626",
        phoneNumber: "714.955.4551",
        faxNumber: "714.966.0663",
      },
      clientInfo: {
        fullName: "Kristine Allen",
        firstName: "Kristine",
        lastName: "Allen",
        pronoun1: "she",
        pronoun2: "her",
        title: "Ms",
      },
      adverseInsuranceAdjuster: {
        title: "Mr",
        pronoun1: "he",
        pronoun2: "him",
        fullName: "David Mulholland",
        firstName: "David",
        lastName: "Mullholland",
        companyName: "State Farm",
        streetAddress: "P.O. Box 106171",
        city: "Atlanta",
        fullStateName: "Georgia",
        zipCode: "30348",
        email: "statefarmclaims@statefarm.com",
        claimNumber: "75-68H9-93X",
      },
    };
    // MAKE SECTION CHILDREN
    this.#sectionChildren.push(
      this.#makeParagraphWithLogo(
        this.#makeFirstPageLogo(
          this.#logoImageType.ext,
          this.#logoBuffer,
          this.#desiredLogoWidthForTitlePage,
          this.#proportionalLogoHeightForTitlePage,
        ),
      ),
    );
    this.#sectionChildren.push(new Paragraph({}));
    this.#sectionChildren.push(
      new Paragraph({
        children: [
          new TextRun(
            new Intl.DateTimeFormat("en-US", this.#dateOptions).format(new Date()),
          ),
        ],
        alignment: AlignmentType.CENTER,
      }),
    );
    this.#sectionChildren.push(new Paragraph({}));
    this.#sectionChildren.push(new Paragraph({}));
    this.#sectionChildren.push(
      new Paragraph({
        text: "PURSUANT TO EVIDENCE CODE §§ 1152 AND 1154",
        style: "CenterBoldItalics",
      }),
    );
    this.#sectionChildren.push(new Paragraph({}));
    this.#sectionChildren.push(new Paragraph({}));
    this.#sectionChildren.push(
      new Paragraph({
        text: `VIA ELECTRONIC MAIL ONLY (${inputVariables.adverseInsuranceAdjuster.email})`,
        style: "LeftBoldItalics",
      }),
    );

    this.#sectionChildren.push(new Paragraph({}));
    this.#sectionChildren.push(new Paragraph(
        `${inputVariables.adverseInsuranceAdjuster.firstName} ${inputVariables.adverseInsuranceAdjuster.lastName}`,
      ));
    this.#sectionChildren.push(
      new Paragraph(`${inputVariables.adverseInsuranceAdjuster.companyName}`)
    );
    this.#sectionChildren.push(
      new Paragraph(`${inputVariables.adverseInsuranceAdjuster.streetAddress}`)
    )
    this.#sectionChildren.push(
      new Paragraph(
        `${inputVariables.adverseInsuranceAdjuster.city}, ${inputVariables.adverseInsuranceAdjuster.fullStateName} ${inputVariables.adverseInsuranceAdjuster.zipCode}`,
      )
    )
    this.#sectionChildren.push(new Paragraph({}));
    this.#sectionChildren.push(
this.#makeReTable(
        inputVariables.clientInfo.fullName,
        new Intl.DateTimeFormat("en-US", this.#dateOptions).format(
          new Date(inputVariables.accidentDetails.dateOfLoss),
        ),
        inputVariables.adverseInsuranceAdjuster.claimNumber,
      )
    )
    this.#sectionChildren.push(new Paragraph({}));
    this.#sectionChildren.push(
      new Paragraph(
        `Dear ${inputVariables.adverseInsuranceAdjuster.title}. ${inputVariables.adverseInsuranceAdjuster.lastName}:`,
      )
    )
    this.#sectionChildren.push(new Paragraph({}));
    this.#sectionChildren.push(
new Paragraph({
        text: `This letter shall serve as a statement of ${inputVariables.clientInfo.fullName}’s damages as a result of the above-referenced loss.  Enclosed for your review, please find a copy of medical and billing records reflecting the treatment ${inputVariables.clientInfo.pronoun1} received as a result of her collision caused by your insured.`,
        style: "Paragraph",
      })
    )
    this.#sectionChildren.push(
this.#makeParagraphWithFootnote(
        "On May 24, 2024, Ms. Allen was the front passenger in a vehicle driving on Alton Parkway in the city of Irvine, state of California.{footnote}  As the vehicle approached the intersection of Alton Parkway and Irvine Boulevard, it stopped for the red light that it was faced it, and, shortly thereafter, it was forcefully struck in the rear by your insured’s vehicle. As a result of the force and nature of the impact, Ms. Allen sustained bodily injuries necessitating medical attention.",
        "{footnote}",
        "To note, Ms. Allen’s husband was the driver of the vehicle, and he is not presenting a bodily injury claim, at least to the best of my knowledge",
      )
    )
    this.#sectionChildren.push(
this.#makeParagraphWithFootnote(
        "After the accident occurred, Ms. Allen began to experience pain and soreness throughout her body which grew progressively worse, so she went to a massage therapist the following day to see if a massage might provide her some relief.{footnote} After realizing that the massage only provided temporary relief and that her pain continued to linger and heighten on occasion, she visited Dr. John Chen at Compassion Chiropractic for an evaluation of his injuries. At the time of her initial evaluation, Ms. Allen complained of dull and tight pain in her neck, dull pain and tightness in her shoulders, and dull and achy pain in her back, all of which she rated at a moderate level.",
        "{footnote}",
        "Ms. Allen continued to receive massages when her pain flared up as indicated in the enclosed receipts, and she will continue to do so for an indefinite period of time.",
      ),
    )
      



    let doc = await this.#buildDocument(
      this.#defaultProperties,
      this.#footnotesObject,
      this.#makeDefaultHeader(
        this.#makeHeaderLogo(
          this.#logoImageType.ext,
          this.#logoBuffer,
          this.#proportionalLogoWidthForHeader,
          this.#desiredLogoHeightForHeader,
        ),
      ),
      this.#makeDefaultFooter(
        inputVariables.attorney.streetAddress,
        inputVariables.attorney.city,
        inputVariables.attorney.fullStateName,
        inputVariables.attorney.zipCode,
        inputVariables.attorney.phoneNumber,
        inputVariables.attorney.faxNumber,
        inputVariables.attorney.email,
        inputVariables.attorney.website,
      ),
      this.#sectionChildren,
    );
    Packer.toBuffer(doc).then((buffer) => {
      fs.writeFileSync("My Document.docx", buffer);
    });
  }

  #makeFirstPageLogo(imageExt, imageBuffer, width, height) {
    console.log("height", height);
    console.log("width", width);
    return new ImageRun({
      type: imageExt,
      data: imageBuffer,
      transformation: {
        width: width,
        height: height,
      },
    });
  }

  #makeHeaderLogo(impageExtension, imageBuffer, width, height) {
    return new ImageRun({
      type: impageExtension,
      data: imageBuffer,
      transformation: {
        width: width,
        height: height,
      },
      floating: {
        horizontalPosition: {
          relative: HorizontalPositionRelativeFrom.RIGHT_MARGIN,
          offset: -(width * 10000),
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
  }

  #makeDefaultFooter(
    streetAddress,
    city,
    state,
    zip,
    phone,
    fax,
    email,
    website,
  ) {
    return {
      first: new Footer({
        children: [
          new Paragraph({
            style: "Footer",
            text: `${streetAddress} ǁ ${city}, ${state} ${zip}`,
          }),
          new Paragraph({
            style: "Footer",
            text: `Phone: ${phone} ǁ Fax: ${fax}`,
          }),
          new Paragraph({
            style: "Footer",
            children: [new TextRun(email)],
          }),
          new Paragraph({
            style: "Footer",
            children: [
              new ExternalHyperlink({
                children: [
                  new TextRun({
                    text: website,
                    style: "Hyperlink",
                    bold: true,
                  }),
                ],
                link: `https://${website}`,
              }),
            ],
          }),
        ],
      }),
    };
  }

  #makeDefaultHeader(headerLogo) {
    return {
      default: new Header({
        children: [
          new Paragraph(
            new Intl.DateTimeFormat("en-US", this.#dateOptions).format(new Date()),
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
            children: [headerLogo],
          }),
        ],
      }),
    };
  }

  #makeReTable(clientFullName, DOL, claimNumber) {
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
                      text: `${clientFullName}’s Automobile Accident Dated ${DOL} (Claim Number: ${claimNumber})`,
                      italics: true,
                      bold: true,
                      underline: true,
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

  #addToFootnoteObject(footnoteText) {
    let nextNumber = 0;
    this.#footnotesObject[
      parseInt(Object.keys(this.#footnotesObject).length + 1)
    ] = {
      children: [
        new Paragraph({
          style: "Footnote",
          children: [new TextRun(" "), new TextRun(footnoteText)],
        }),
      ],
    };
    return Object.keys(this.#footnotesObject).length;
  }

  #makeParagraphWithLogo(firstPageLogo) {
    return new Paragraph({
      children: [firstPageLogo],
      alignment: AlignmentType.CENTER,
    });
  }

  #makeParagraphWithFootnote(wholeParagraph, footnoteDelimieter, footnoteText) {
    let footnoteIndex = this.#addToFootnoteObject(footnoteText);
    let sentences = wholeParagraph
      .split(footnoteDelimieter)
      .map((sentences) => sentences.trim());

    let retVal = new Paragraph({
      children: [
        new TextRun({
          text: sentences[0],
        }),
        new FootnoteReferenceRun(footnoteIndex),
        new TextRun("  "),
        new TextRun(sentences[1]),
      ],
      style: "Paragraph",
    });
    return retVal;
  }

  // #makeSectionChildren(inputVariables) {
  //   let retVal = [
  //     this.#makeParagraphWithLogo(
  //       this.#makeFirstPageLogo(
  //         this.#logoImageType.ext,
  //         this.#logoBuffer,
  //         this.#desiredLogoWidthForTitlePage,
  //         this.#proportionalLogoHeightForTitlePage,
  //       ),
  //     ),
  //     new Paragraph({}),
  //     new Paragraph({
  //       children: [
  //         new TextRun(
  //           new Intl.DateTimeFormat("en-US", dateOptions).format(new Date()),
  //         ),
  //       ],
  //       alignment: AlignmentType.CENTER,
  //     }),
  //     new Paragraph({}),
  //     new Paragraph({}),
  //     new Paragraph({
  //       text: "PURSUANT TO EVIDENCE CODE §§ 1152 AND 1154",
  //       style: "CenterBoldItalics",
  //     }),
  //     new Paragraph({}),
  //     new Paragraph({}),
  //     new Paragraph({
  //       text: `VIA ELECTRONIC MAIL ONLY (${inputVariables.adverseInsuranceAdjuster.email})`,
  //       style: "LeftBoldItalics",
  //     }),
  //     new Paragraph({}),
  //     new Paragraph(
  //       `${inputVariables.adverseInsuranceAdjuster.firstName} ${inputVariables.adverseInsuranceAdjuster.lastName}`,
  //     ),
  //     new Paragraph(`${inputVariables.adverseInsuranceAdjuster.companyName}`),
  //     new Paragraph(`${inputVariables.adverseInsuranceAdjuster.streetAddress}`),
  //     new Paragraph(
  //       `${inputVariables.adverseInsuranceAdjuster.city}, ${inputVariables.adverseInsuranceAdjuster.fullStateName} ${inputVariables.adverseInsuranceAdjuster.zipCode}`,
  //     ),
  //     new Paragraph({}),
  //     this.#makeReTable(
  //       inputVariables.clientInfo.fullName,
  //       new Intl.DateTimeFormat("en-US", dateOptions).format(
  //         new Date(inputVariables.accidentDetails.dateOfLoss),
  //       ),
  //       inputVariables.adverseInsuranceAdjuster.claimNumber,
  //     ),
  //     new Paragraph({}),
  //     new Paragraph(
  //       `Dear ${inputVariables.adverseInsuranceAdjuster.title}. ${inputVariables.adverseInsuranceAdjuster.lastName}:`,
  //     ),
  //     new Paragraph({}),
  //     new Paragraph({
  //       text: `This letter shall serve as a statement of ${inputVariables.clientInfo.fullName}’s damages as a result of the above-referenced loss.  Enclosed for your review, please find a copy of medical and billing records reflecting the treatment ${inputVariables.clientInfo.pronoun1} received as a result of her collision caused by your insured.`,
  //       style: "Paragraph",
  //     }),
  //     this.#makeParagraphWithFootnote(
  //       "On May 24, 2024, Ms. Allen was the front passenger in a vehicle driving on Alton Parkway in the city of Irvine, state of California.{footnote}  As the vehicle approached the intersection of Alton Parkway and Irvine Boulevard, it stopped for the red light that it was faced it, and, shortly thereafter, it was forcefully struck in the rear by your insured’s vehicle. As a result of the force and nature of the impact, Ms. Allen sustained bodily injuries necessitating medical attention.",
  //       "{footnote}",
  //       "To note, Ms. Allen’s husband was the driver of the vehicle, and he is not presenting a bodily injury claim, at least to the best of my knowledge",
  //     ),
  //     this.#makeParagraphWithFootnote(
  //       "After the accident occurred, Ms. Allen began to experience pain and soreness throughout her body which grew progressively worse, so she went to a massage therapist the following day to see if a massage might provide her some relief.{footnote} After realizing that the massage only provided temporary relief and that her pain continued to linger and heighten on occasion, she visited Dr. John Chen at Compassion Chiropractic for an evaluation of his injuries. At the time of her initial evaluation, Ms. Allen complained of dull and tight pain in her neck, dull pain and tightness in her shoulders, and dull and achy pain in her back, all of which she rated at a moderate level.",
  //       "{footnote}",
  //       "Ms. Allen continued to receive massages when her pain flared up as indicated in the enclosed receipts, and she will continue to do so for an indefinite period of time.",
  //     ),
  //   ];
  //   retVal = [];
  //   return this.#sectionChildren;
  // }

  async #buildDocument(
    properties,
    footNotesObject,
    headers,
    footers,
    sectionChildren,
  ) {

    let retVal = new Document({
      footnotes: footNotesObject,
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
                      hanging: convertInchesToTwip(0.5),
                    },
                    spacing: {
                      after: 300,
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
              font: "Times New Roman",
            },
          },
        },
        paragraphStyles: [
          {
            name: "CenterBoldItalics",
            basedOn: "Normal",
            next: "Normal",
            quickFormat: true,
            paragraph: {
              // spacing: {
              //     after: 240
              // },
              // indent: {
              //     firstLine: 720
              // },
              alignment: AlignmentType.CENTER,
            },
            run: {
              bold: true,
              italics: true,
              underline: true,
            },
          },
          {
            name: "LeftBoldItalics",
            basedOn: "Normal",
            next: "Normal",
            quickFormat: true,
            paragraph: {
              // spacing: {
              //     after: 240
              // },
              // indent: {
              //     firstLine: 720
              // },
              alignment: AlignmentType.LEFT,
            },
            run: {
              bold: true,
              italics: true,
              underline: true,
            },
          },
          {
            name: "Paragraph",
            basedOn: "Normal",
            next: "Normal",
            quickFormat: true,
            paragraph: {
              spacing: {
                after: 300,
              },
              indent: {
                firstLine: 720,
              },
              alignment: AlignmentType.JUSTIFIED,
            },
          },
          {
            name: "Footnote",
            paragraph: {
              spacing: {
                after: 240,
              },
              alignment: AlignmentType.LEFT,
            },
            run: {
              font: "Calibri",
              size: "10pt",
            },
          },
          {
            name: "Encl",
            paragraph: {
              spacing: {
                after: 240,
              },
              alignment: AlignmentType.LEFT,
            },
            run: {
              size: "10pt",
            },
          },
          {
            name: "Footer",
            basedOn: "Normal",
            next: "Normal",
            quickFormat: true,
            paragraph: {
              alignment: AlignmentType.CENTER,
            },
            run: {
              size: "10pt",
            },
          },
        ],
      },
      sections: [
        {
          properties: properties,
          footers: footers,
          headers: headers,
          children: sectionChildren,
        },
      ],
    });
    return retVal;
  }

  calcSizeFromInches(inches) {
    return Math.round(inches * (275 / 1.16));
  }

  sayHello() {
    return `Hello, my name is ${this.#name}!`;
  }
}
