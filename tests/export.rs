use assert_cmd::prelude::*;
use std::process::Command;
use tempfile::tempdir;
use std::fs;
use std::io::Write;

#[test]
fn test_export_command() {
    let dir = tempdir().unwrap();
    let out_dir = dir.path().join("output");
    let mut cmd = Command::new(env!("CARGO_BIN_EXE_xcsv"));

    // Create a dummy xlsx file
    let xlsx_path = dir.path().join("input.xlsx");
    let mut zip = zip::ZipWriter::new(fs::File::create(&xlsx_path).unwrap());
    let options = zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Stored);

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>"#).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="s"><v>0</v></c>
            <c r="B1" t="s"><v>1</v></c>
        </row>
        <row r="2">
            <c r="A2"><v>10.123</v></c>
            <c r="B2"><v>-20.456</v></c>
        </row>
    </sheetData>
</worksheet>"#).unwrap();

    zip.start_file("xl/sharedStrings.xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
    <si><t>origin_latitude</t></si>
    <si><t>origin_longitude</t></si>
</sst>"#).unwrap();

    zip.finish().unwrap();

    cmd.arg(&xlsx_path).arg("export").arg("--out-dir").arg(&out_dir);
    cmd.assert().success();

    let csv_path = out_dir.join("sheet1.csv");
    let csv_content = fs::read_to_string(csv_path).unwrap();
    let expected_content = "origin_latitude,origin_longitude\n10.123,-20.456\n";
    assert_eq!(csv_content, expected_content);
}
