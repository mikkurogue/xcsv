use assert_cmd::prelude::*;
use std::process::Command;
use tempfile::tempdir;
use std::fs;
use std::io::Write;

#[test]
fn test_list_command() {
    let dir = tempdir().unwrap();
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

    zip.finish().unwrap();

    cmd.arg(&xlsx_path).arg("list");
    cmd.assert().success().stdout("Sheet1\n");
}
