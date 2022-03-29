# pdf_osiris

Parses Osiris pdf files

## Getting Started

These instructions will show some examples to run the script

### Prerequisites

pdfplumber, xlsxwriter and tabulate need to be installed on your system


### Installing

Use pip to install pdfplumber, xlsxwriter and tabulate.
- pdfplumber:

```
pip install pdfplumber
```

- xlsxwriter

```
pip install xlsxwriter
```

- tabulate

```
pip install tabulate
```

## Running the parse_pdf.py module


Example:
```
python parse_pdf.py -v my_pdf.pdf my_excel
```

### Arguments:

Required argument:
- my_pdf.pdf: the path to the pdf file
- my_excel: the path to the Excel file. Note: the suffix .xlsx is NOT required. 

Optional arguments:
- v: --verbose mode

## Running the batch_parse_pdf.py module

This script will convert each pdf into an Excel file.
Safe all pdf files in a pdf folder that is located in the folder where batch_parse_pdf.py resides.

Example:
```
python batch_parse_pdf.py -v pdf_folder excel_folder
```
### Arguments:

Required argument:
- pdf_folder: the path to the folder where the pdf files are located  
- excel_folder: the path to the folder where the pdf files are located  

Optional arguments:
- v: --verbose mode

## Authors

- *Jurre Hageman*


## License

This project is licensed under the GNU General Public License (GPL)
