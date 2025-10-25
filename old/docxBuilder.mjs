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

