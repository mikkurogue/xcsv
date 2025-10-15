# libxcsv

A lightweight, low-level library for converting Excel (.xlsx) sheets to CSV by stream-parsing the underlying XML.

## Usage

`libxcsv` is the core engine for the `xcsv` CLI tool. It operates by:

1.  Opening the `.xlsx` file as a zip archive (`open_zip`).
2.  Parsing metadata from `workbook.xml` (`parse_workbook`), `sharedStrings.xml` (`read_shared_strings`), and `styles.xml` (`parse_styles`).
3.  Streaming the contents of each worksheet and converting rows to CSV format using `export_sheet_xml_to_csv`.

## Core Functions

-   `open_zip()`: Opens the `.xlsx` file.
-   `parse_workbook()` & `parse_workbook_rels()`: Reads sheet metadata.
-   `read_shared_strings()`: Parses the shared string table.
-   `parse_styles()`: Parses cell styles for date/time formatting.
-   `export_sheet_xml_to_csv()`: The main function to convert a sheet XML to a CSV file.