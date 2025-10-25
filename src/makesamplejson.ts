import * as path from 'path';
import * as fs from 'fs';
import { fileURLToPath } from 'url';
import { Command } from 'commander';
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function replaceString(originalText: string, oldString: string, newString: string) {
    let newText = originalText.replace(new RegExp(`\\b${oldString}\\b`, 'g'), newString);
    return newText;
}

const program = new Command();
program
    .name('make-sample-docx')
    .description('hello world')
    .version('1.0.0')
    .option('-d, --data <path to datafile>', 'Define the path to the data file')
    .option('-l, --logo <path to image file for logo>', 'Provide the path to the image file to be used for the logo');

program.parse(process.argv);

const options = program.opts();

const filePath = path.join(__dirname, '../', options.data);

let jsonString = fs.readFileSync(filePath, 'utf8');
jsonString = replaceString(jsonString, 'SimonMed Southern CA Laguna Hills','The Anvil-Free Imaging Center');

jsonString = replaceString(jsonString, 'Kristine Allen','Elmer Fudd');


jsonString = replaceString(jsonString, 'Ms. Allen','Mr. Fudd');
jsonString = replaceString(jsonString, 'her husband','his wife');

jsonString = replaceString(jsonString, 'Kristine','Elmer');

jsonString = replaceString(jsonString, 'Allen','Fudd');


jsonString = replaceString(jsonString, 'husband','wife');


jsonString = replaceString(jsonString, 'but he','but she');


jsonString = replaceString(jsonString, 'diagnosed her with the following','diagnosed him with the following');

jsonString = replaceString(jsonString, 'putting her in a constant fight','putting him in a constant fight');



jsonString = replaceString(jsonString, 'especially with her daughter who wanted to spend more time with her mother','especially with his daughter who wanted to spend more time with her father');


jsonString = replaceString(jsonString, 'forcing her to spend extra','forcing him to spend extra');


jsonString = replaceString(jsonString, 'Superior Imaging','The Anvil-Free Imaging Center');


jsonString = replaceString(jsonString, 'Regenerative Institute Of Newport Beach','The TNT Triage & Numbness Center of Agua Dulce');


jsonString = replaceString(jsonString, 'she','he');
jsonString = replaceString(jsonString, 'She','He');

jsonString = replaceString(jsonString, 'Jasmine Wellness Spa',"Moe N' Larry's Wellness Spa");

jsonString = replaceString(jsonString, 'Dr. John Chi-Chang Chen','Dr. Boneaparte (Bones) Cracker');
jsonString = replaceString(jsonString, 'Dr. Chen','Dr. Cracker');

jsonString = replaceString(jsonString, 'Compassion Chiropractic','Acme Spinal Adjustments Center');

jsonString = replaceString(jsonString, 'Dr. Khyber Zaffarkhan',"Dr. Percival (Patch) Palliative");
jsonString = replaceString(jsonString, 'Dr. Zaffarkhan',"Dr. Palliative");
jsonString = replaceString(jsonString, 'Regenerative Institute of Newport Beach',"The TNT Triage & Numbness Center of Agua Dulce");

jsonString = replaceString(jsonString, 'city of Irvine','city of Agua Dulce');


jsonString = replaceString(jsonString, 'Alton','Dust Devil');


jsonString = replaceString(jsonString, 'Irvine Boulevard','Saguaro Junction Highway');


jsonString = replaceString(jsonString, 'Samir I. Sheth','Wile E. Coyote');
jsonString = replaceString(jsonString, 'SAMIR','WILE');

jsonString = replaceString(jsonString, 'samir@sheth-law.com','wile.e.coyote@acme-gizmo-liability.com');

jsonString = replaceString(jsonString, 'www.sheth-law.com','www.acme-gizmo-liability.com');

jsonString = replaceString(jsonString, '650 Town Center Drive, Suite 1400','P.O. Box 140 (Explosives Division)');

jsonString = replaceString(jsonString, 'Costa Mesa','Agua Dulce');

jsonString = replaceString(jsonString, '92626','91390');

jsonString = replaceString(jsonString, '714.955.4551','928.555.1234');

jsonString = replaceString(jsonString, '714.966.0663','928.555.9876');

jsonString = replaceString(jsonString, 'David Mulholland','Brock (The Block) Hardcastle');
jsonString = replaceString(jsonString, 'Mulholland','Hardcastle');
jsonString = replaceString(jsonString, 'State Farm','C. Y. A. Catastrophe & Annuity, Inc.');
jsonString = replaceString(jsonString, 'Mulholland','Hardcastle');


jsonString = replaceString(jsonString, 'David','Brock (The Block)');

jsonString = replaceString(jsonString, 'P.O. Box 106171','P.O. Box 99 (Zero-Payout Division)');

jsonString = replaceString(jsonString, 'Atlanta','Denial Heights');
jsonString = replaceString(jsonString, 'Georgia','Delaware');
jsonString = replaceString(jsonString, '30348','19800');
jsonString = replaceString(jsonString, 'statefarmclaims@statefarm.com','claims@cya-denies.com');
jsonString = replaceString(jsonString, '75-68H9-93X','FUDD-FAIL-404-NO-PAY');




const outputFilePath = path.join(__dirname,'../','sampledata2.json');

fs.writeFileSync(outputFilePath, jsonString);

