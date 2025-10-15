## xcsv
![build](https://github.com/mikkurogue/xcsv/actions/workflows/rust.yml/badge.svg)
Convert Excel (.xlsx) workbooks to CSV, one CSV per sheet.

### Features

- Lowercase filenames derived from sheet names (non-alphanumerics replaced with `_`)
- Streams worksheet XML to CSV using `quick-xml` and `csv` (low memory usage)
- Preserves empty cells as empty CSV fields (no padding heuristics needed)
- **Excel Date/Time Support**: Converts Excel serial dates to ISO 8601 format
- **Comprehensive Cell Types**: Shared strings, booleans, inline strings, formulas, error values, and numbers
- **Smart Date Detection**: Automatically detects and converts Excel date serial numbers

### Install / Build

Note; building from source can result in unintended behaviour I recommend to always follow the installation steps found on each release.

Build from source:

```bash
cargo build --release
```

Binary path:

```bash
./target/release/xcsv
```

Installation with `cargo install`

```bash
cargo install --path . --locked --force
```

Install release;
Follow the instructions at [releases](https://github.com/mikkurogue/xcsv/releases)

### Usage

Show help:

```bash
xcsv --help
```

#### List sheets

Print sheet names in an `.xlsx` file:

```bash
xcsv <path-to-file.xlsx> list
```

Example:

```bash
xcsv input.xlsx list
```

#### Export all sheets to CSV

Export each sheet to its own CSV file in the output directory. Filenames are the sheet names lowercased with non-alphanumeric characters mapped to `_` and a `.csv` suffix.

```bash
xcsv <path-to-file.xlsx> export --out <output-dir>
# or
xcsv <path-to-file.xlsx> export -o <output-dir>
```

**CSV Delimiter Options:**

```bash
# Default: comma delimiter
xcsv input.xlsx export -o out

# Semicolon delimiter (useful for European locales)
xcsv input.xlsx export -o out --delimiter ";"
# or
xcsv input.xlsx export -o out -d ";"
```

Examples:

```bash
# Standard comma-separated CSV
xcsv input.xlsx export -o out
# writes files like:
# out/sheet1.csv
# out/sheet2.csv
# out/financial_sheet_and_stuff_top_secret.csv

# Semicolon-separated CSV (European format)
xcsv input.xlsx export -o out --delimiter ";"
# writes files with semicolon delimiters instead of commas
```

### Notes and behavior

- **Memory Efficient**: Streams XML directly from ZIP entries without loading entire files into memory
- **Flexible CSV Output**: Row lengths may vary across the file if trailing empty cells are omitted by Excel
- **Empty Cell Preservation**: Empty cells are preserved as empty fields within a row based on cell coordinates
- **Excel Date Conversion**: Automatically converts Excel serial dates (e.g., `44927.0` → `2023-01-01T00:00:00.000Z`)
- **Supported Cell Types**:
  - Shared strings (`t="s"`) - References to shared string table
  - Inline strings (`t="inlineStr"`) - Direct text content
  - Booleans (`t="b"`) - TRUE/FALSE values
  - Formula results (`t="str"`) - String results from formulas
  - Error values (`t="e"`) - Excel error codes like #N/A, #VALUE!
  - Numeric values - With intelligent date detection
- **CSV Delimiter Support**: Choose between comma (`,`) and semicolon (`;`) delimiters

### Limitations / roadmap

- Only `.xlsx` (Office Open XML) files are supported. Legacy `.xls` is not supported.
- Date detection is heuristic-based (numbers ≥1000 or with fractional parts in reasonable date range)
- Number format styles from Excel are not preserved (dates converted to ISO format, not original formatting)
- Future options that could be added:
  - Select specific sheets to export
  - Custom CSV quote/escape characters
  - Normalize row lengths to max columns (pad trailing empties)
  - Preserve Excel number formatting
  - Parse number format codes for more accurate date detection

### License

MIT


