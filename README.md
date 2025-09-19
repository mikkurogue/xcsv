## xcsv

Convert Excel (.xlsx) workbooks to CSV, one CSV per sheet.

### Features

- Lowercase filenames derived from sheet names (non-alphanumerics replaced with `_`)
- Streams worksheet XML to CSV using `quick-xml` and `csv`
- Preserves empty cells as empty CSV fields (no padding heuristics needed)
- Handles shared strings, booleans, inline strings, and raw numbers

### Install / Build

Build release binary:

```bash
cargo build --release
```

Binary path:

```bash
./target/release/xcsv
```

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

Examples:

```bash
xcsv input.xlsx export -o out
# writes files like:
# out/sheet1.csv
# out/emission_assets.csv
# out/shipment_definitions.csv
```

### Notes and behavior

- Output CSV rows are written as encountered; row lengths may vary across the file if trailing empty cells are omitted by Excel. The writer is configured as flexible to allow this.
- Empty cells are preserved as empty fields within a row based on cell coordinates.
- Supported cell types: shared strings (`t="s"`), inline strings, booleans, and raw numeric/text values. Date/time formatting from Excel number formats is not applied (values are exported as-is).

### Limitations / roadmap

- Date/time and formatted numbers are not converted; they are exported as raw values.
- Only `.xlsx` (Office Open XML) files are supported. Legacy `.xls` is not supported.
- Future options that could be added:
  - Select specific sheets to export
  - Custom CSV delimiter/quote/escape
  - Normalize row lengths to max columns (pad trailing empties)

### License

MIT


