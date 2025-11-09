export interface Footnote {
    children: Array<any>
}

export interface ImageDimensions {
    height: number,
    width: number
}

export type ImageType = "jpg" | "png" | "gif" | "bmp" | "svg" | "webp";

export interface AccidentData {
    letterDetails: {
        dateOfLoss: Date,
        formattedResponseByDate: string,
        formattedDateOfLetter: string
    },
    attorney: {
        attorneyFullName: string,
        email: string,
        website: string,
        streetAddress: string,
        city: string,
        stateCode: string,
        fullStateName: string,
        zipCode: string,
        phoneNumber: string,
        faxNumber?: string
    },
    clientInfo: {
        fullName: string,
        firstName: string,
        lastName: string,
        pronoun1: string,
        pronoun2: string,
        title: string
    },
    adverseInsuranceAdjusterInfo: {
        title: string,
        pronoun1: string,
        pronoun2: string,
        fullName: string,
        firstName: string,
        lastName: string,
        companyName: string,
        streetAddress: string,
        city: string,
        stateCode: string,
        fullStateName: string,
        zipCode: string,
        email: string,
        claimNumber: string
    },
    paragraphs: [
        {
            text: string,
            footnote?: string,
            type: "bullet" | "paragraph" | "bulletBold" | "pageBreak" | "encl"
            bulletInstance?: number
        }
    ]
}