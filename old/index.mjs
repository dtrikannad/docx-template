    import JSZip from 'jszip';
    import Docxtemplater from 'docxtemplater';
import fs from 'fs';
import { promises as fsPromises } from 'fs';
    import path from 'path';
    import { fileURLToPath } from 'url';
import yargs from 'yargs';
import { hideBin } from 'yargs/helpers';
import SubtemplateModule from 'docxtemplater-subtemplate-module';
import FootnoteModule from 'docxtemplater-footnote-module';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const argv = yargs(hideBin(process.argv))
  .usage('Usage: $0 -t <templateFile> -d <dataFile> -o <outputFile>')
  .option('t', {
    alias: 'templateFile',
    describe: 'Path to the template file',
    demandOption: true,
    type: 'string'
  })
  .option('d', {
    alias: 'dataFile',
    describe: 'Path to the data file',
    demandOption: true,
    type: 'string'
  })
  .option('o', {
    alias: 'outputFile',
    describe: 'Path to the output file',
    demandOption: true,
    type: 'string'
  })
  .help()
  .argv;

    function validateFilesExist(filePaths) {
        filePaths.forEach(filePath => {
            if (!fs.existsSync(filePath)) {
                console.error(`Error: File '${filePath}' does not exist`);
                process.exit(1);
            }
        });
    }
    function processFiles(dataPath, templatePath, outputPath) {
            try {
                // Example: Read data and template files
                const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
                
        // Load the template
        const content = fs.readFileSync(path.resolve(__dirname, templatePath), 'binary');
    
        const zip = new JSZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true, // Crucial for handling list items correctly
            linebreaks: true, // Optional: for preserving line breaks in text
        });
        
        try {
            // Render the document
            doc.render(data);
        } catch (error) {
            // Handle errors during rendering
            console.error('Error rendering document:', error);
            throw error;
        }
    
        // Generate the output DOCX
        const buf = doc.getZip().generate({
            type: 'nodebuffer',
            compression: 'DEFLATE',
        });
    
        // Save the output file
        fs.writeFileSync(path.resolve(__dirname, outputPath), buf);
    
        console.log('Document generated successfully!');
                
            } catch (error) {
                console.error('Error processing files:', error.message);
                process.exit(1);
            }
        }

        function main() {
          const dataFilePath = argv.dataFile;
          const templateFilePath = argv.templateFile;
          const outputFilePath = argv.outputFile;
          
          // Validate all required files exist
          validateFilesExist([dataFilePath, templateFilePath]);
          
          console.log('Data file:', dataFilePath);
          console.log('Template file:', templateFilePath);
          console.log('Output file:', outputFilePath);
          
          // Your processing logic here
          processFiles(dataFilePath, templateFilePath, outputFilePath);
      }
main();

// Your logic here