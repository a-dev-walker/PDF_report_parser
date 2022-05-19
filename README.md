# PDF_report_parser
Compress medical reports to more easily extract charts and other wanted information

This code is a part of a pipeline to process genetics/medical reports to be more easily ingested into databases. 

The reports are initially exported from their respective softwares before being converted to .xlsx files by Adobe Acrobat Pro XI using it's batch conversion macros. 

In order to use the conversion macros, the user needs to install the CHLA version of Adobe Pro XI. They must then go to the "tools" menu. Click "create new action" and select "save and export" from the left menu and add a "save" module to your new macro. Using the "specify settings" button under the added save module, select "export file(s) to alternative output" and select excel workbook. You can then save and use the macro to batch convert folders of pdf's to excel files.

Once the files are in .xlsx format, they are able to be processed with the code in this repository and output .csv files with more easily extractable tables and other information. 

The manual code can be used for individual files where the user pastes the file location of the desired to-be-converted .xlsx file. 

The batch code (which can also run individual files) takes no arguments and when run will ask the user which folder they would like to select for conversion. The software only pays attention to files of.xlsx extensions and creates a subfolder for placing the processed data. 

This is very versatile code and can be used on a wide swath of reports to condense them down into more easily ingestible bits and pieces.
