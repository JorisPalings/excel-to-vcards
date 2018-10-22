#!/usr/bin/env node

const program = require('commander');
const path = require('path');
const fs = require('fs');
const stream = require('stream');
const csvParse = require('csv-parse');
const xlsx = require('xlsx');

class Contact {
  /**
   * Represents a single contact.
   * @constructor
   * @param {string} firstName - The first name of the contact.
   * @param {string} lastName - The last name of the contact.
   * @param {string} emailAddress - The email address of the contact.
   * @param {string} telephoneNumber - The telephone number of the contact.
   */
  constructor(firstName, lastName, emailAddress, telephoneNumber) {
    this.firstName = firstName;
    this.lastName = lastName;
    this.emailAddress = emailAddress;
    this.telephoneNumber = telephoneNumber;
  }
}

program
  .version('1.0.0', '-v, -V, --version')
  .option('-i, --input <path>', 'path to the .csv input file')
  .option('-d, --delimiter [delimiter]', 'delimiter used in the .csv input file')
  .option('-o, --output [directory]', 'output directory for the .vcf file (defaults to current directory)')
  .option('-s, --start [row]', '1-based index of the first data row (defaults to first row)')
  .option('-e, --end [row]', '1-based index of the last data row (defaults to last row with data)')
  .option('-t, --telephone', 'whether or not the telephone number should be formatted')
  .parse(process.argv);

(async () => {
  if (program.input) {
    if (isInputFileReadable(program.input)) {
      const contacts = await parseInputFileToContacts(program.input);
      let contactsAsVcards = '';
      contacts.forEach(contact => {
        contactsAsVcards += convertContactToVcard(contact);
      });
      fs.writeFileSync(createOutputFilePath(), contactsAsVcards);
      console.info(`Successfully converted ${contacts.length} rows to contacts.`);
    }
  } else {
    console.error(`Don't forget the -i or --input argument to specify a file to convert!`);
  }
})();

/**
 * Test the user's permissions for read access to a file.
 * @param {string} inputFile - The path to the .csv, .xls or .xlsx input file.
 * @returns {boolean} Whether the user has the permissions for read access to the .csv, .xls or .xlsx input file.
 */
function isInputFileReadable(inputFile) {
  try {
    fs.accessSync(inputFile, fs.constants.R_OK);
    return true;
  } catch (error) {
    console.error(`Input file "${inputFile}" is not readable:`, error);
    return false;
  }
}

/**
 * Get the extension of a file. 
 * @param {string} file - A file path.
 * @returns {string} The file's extension.
 */
function getFileExtension(file) {
  try {
    const fileExtension = path.extname(file);
    return fileExtension;
  } catch(error) {
    console.error(`Unable to determine the extension of "${file}":`, error);
  }
}

/**
 * Get the basename of a file.
 * @param {string} file = A file path.
 * @returns {string} The basename of the file.
 */
function getFileBaseName(file) {
  try {
    const fileBaseName = path.basename(file).split('.')[0];
    return fileBaseName;
  } catch(error) {
    console.error(`Unable to determine the extension of "${file}":`, error);
  }
}

/**
 * Construct the path of the .vcf vCard output file.
 * @returns {string} The path to the .vcf vCard output file.
 */
function createOutputFilePath() {
  const outputDirectory = program.output || './';
  const outputFileName = getFileBaseName(program.input);
  const outputFilePath = `${path.join(outputDirectory, outputFileName)}.vcf`;
  return outputFilePath;
}

/**
 * Parse the input file to an Array of Contacts based on its extension.
 * @param {string} inputFile - The path to a .csv, .xls or .xlsx file.
 * @returns {Array} An Array of Contacts parsed from the .csv, .sls or .xlsx file.
 */
async function parseInputFileToContacts(inputFile) {
  const inputFileExtension = getFileExtension(inputFile);
  switch(inputFileExtension) {
    case '.csv':
      const contactsFromCsv = await parseCsvFileToContacts(inputFile);
      return contactsFromCsv;
    case '.xlsx':
    case '.xls':
      const contactsFromXlsx = await parseXlsxFileToContacts(inputFile);
      return contactsFromXlsx;
    default:
      console.error(`Unable to parse "${inputFile}": "${inputFileExtension}" is not a supported format`);
  }
}

/**
 * Parse the input .csv file to an Array of Contacts.
 * @param {string} inputFile - The path to a .csv file.
 * @returns {Array} An Array of Contacts parsed from the .csv file.
 */
async function parseCsvFileToContacts(inputFile) {
  return new Promise(async (resolve, reject) => {
    try {
      const csvDataStream = fs.createReadStream(inputFile);
      const contacts = await parseCsvDataStreamToContacts(csvDataStream, ';');
      resolve(contacts);
    } catch(error) {
      reject(`Unable to parse input .csv file:`, error);
    }
  });
}

/**
 * Parse the input .xls or .xlsx file to an Array of Contacts.
 * @param {string} inputFile - The path to a .xls or .xlsx file.
 * @returns {Array} An Array of Contacts parsed from the .xls or .xlsx file.
 */
async function parseXlsxFileToContacts(inputFile) {
  return new Promise((resolve, reject) => {
    try {
      const buffers = [];
      fs.createReadStream(inputFile)
        .on('data', data => {
          buffers.push(data);
        })
        .on('error', error => {
          console.error(`Unable to read file "${inputFile}":`, error);
        })
        .on('end', async () => {
          const buffer = Buffer.concat(buffers);
          const workbook = xlsx.read(buffer, { type: 'buffer' });
          const firstWorksheet = workbook.Sheets[workbook.SheetNames[0]];
          // convert the entire worksheet to a single csv string
          const csvData = xlsx.utils.sheet_to_csv(firstWorksheet, { raw: true });
          // convert the csv string into a stream
          const csvDataStream = new stream.Readable;
          csvDataStream.push(csvData);
          csvDataStream.push(null);
  
          const contacts = await parseCsvDataStreamToContacts(csvDataStream, ',');
          resolve(contacts);
        });
    } catch(error) {
      reject(`Unable to parse input .xls(x) file:`, error);
    }
  });
}

/**
 * Parse a CSV data stream of values delimited by the given delimiter to a Stream of Contacts.
 * @param {Object} csvDataStream - A readable Stream of CSV data.
 * @param {string} delimiter - The delimiter used to delimit the data in the CSV data stream.
 * @returns {Array} An Array of Contacts parsed from the CSV data stream.
 */
async function parseCsvDataStreamToContacts(csvDataStream, delimiter) {
  return new Promise((resolve, reject) => {
    try {
      let contacts = [];
      csvDataStream
        .pipe(csvParse({ delimiter }))
        .on('data', contact => {
          contacts.push(new Contact(...contact));
        })
        .on('error', error => {
          console.error(`Unable to read file "${program.input}":`, error);
        })
        .on('end', () => {
          contacts = stripUnwantedRows(contacts);
          resolve(contacts);
        });
    } catch(error) {
      reject(`Unable to parse CSV data stream to contacts:`, error);
    }
  });
}

/**
 * Strips the unwanted rows from the beginning and end of the CSV data.
 * @param {Array} rows - All CSV data rows.
 * @returns {Array} The CSV data rows from the desired start row to the desired end row.
 */
function stripUnwantedRows(rows) {
  const startRow = program.start ? program.start - 1 : 0; // shift from 1-based to 0-based
  const endRow = program.end;
  let wantedRows = [];
  if (endRow) {
    wantedRows = rows.slice(startRow, endRow);
  } else {
    wantedRows = rows.slice(startRow);
  }
  return wantedRows;
}

/**
 * Format a telephone number.
 * @param {string} telephoneNumber - A telephone number to format.
 * @returns {string} The telephone number, formatted.
 */
function formatTelephoneNumber(telephoneNumber) {
  if (program.telephone) {
    try {
      const countryCode = `+${telephoneNumber.substring(0, 2)}`;
      const zoneNumber = `${telephoneNumber.substring(2, 5)}`;
      const subscriberNumber = telephoneNumber.substring(5).replace(/(.{2})/g, '$1 ')
      const formattedTelephoneNumber = `${countryCode} ${zoneNumber} ${subscriberNumber}`;
      return formattedTelephoneNumber;
    } catch(error) {
      console.error('Unable to format telephone number:', error);
    }
  } else {
    return telephoneNumber;
  }
}

/**
 * Formats the current date as a shortened, normalized ISO string to be used as the value of the REV field of a vCard.
 * @returns {string} The current date, formatted as a shortened, normalized ISO string.
 */
function formatCurrentDateAsISO() {
  try {
    const currentDate = new Date();
    const currentIsoDate = currentDate.toISOString();
    const normalizedIsoDate = currentIsoDate.replace(/([-:]|\.\d*)+/g, '');
    return normalizedIsoDate;
  } catch(error) {
    console.error('Unable to format current ISO date:', error);
  }
}

/**
 * 
 * @param {Contact} contact - A Contact to convert to a vCard string.
 * @returns {string} The contact converted to a vCard string.
 */
function convertContactToVcard(contact) {
  let vcardString = 'BEGIN:VCARD\n' + 
    'VERSION:4.0\n';
  if (contact.firstName || contact.lastName) {
    vcardString += `N:${contact.lastName ? contact.lastName : ''};${contact.firstName ? contact.firstName : ''};;;\n` +
    `FN:${contact.firstName} ${contact.lastName}\n`;
  }
  if (contact.telephoneNumber) {
    vcardString += `TEL;TYPE=cell:${formatTelephoneNumber(contact.telephoneNumber)}\n`;
  }
  if (contact.emailAddress) {
    vcardString += `EMAIL:${contact.emailAddress}\n`;
  }
  vcardString += `REV:${formatCurrentDateAsISO()}\n` + 
    'END:VCARD\n';
  return vcardString
}