# Excel to vCards converter

Command-line tool to easily convert Excel files (.csv, .xls, .xlsx) to vCards (.vcf)

## Prerequisites

[npm and Node.js](https://nodejs.org/en/) are required to install and run the script.

## Installation

```
npm install
```

### Usage

```
node index [options]

Options:
    -v, -V, --version             output the version number
    -i, --input <path>            path to the .csv input file
    -d, --delimiter [delimiter]   delimiter used in the .csv input file
    -o, --output [directory]      output directory for the .vcf file (defaults to current directory)
    -s, --start [row]             1-based index of the first data row (defaults to first row)
    -e, --end [row]               1-based index of the last data row (defaults to last row with data)
    -t, --telephone               whether or not the telephone number should be formatted
    -h, --help                    output usage information
```

Only the `-i` or `--input` option is required.

## Built with

* [Node.js](https://nodejs.org/en/) - JavaScript runtime
* [node-csv-parse](https://github.com/adaltas/node-csv-parse) - .csv parser
* [js-xlsx](https://github.com/sheetjs/js-xlsx) - .xls and .xlsx parser

## License

This project is licensed under the GNU GPL v3 License - see the [LICENSE](LICENSE) file for details
