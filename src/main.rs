use std::collections::BTreeMap;
use std::fs::File;
use std::io::{BufRead, BufReader};
use std::path::{Path, PathBuf};

use clap::{Parser, Subcommand};
use anyhow::{Context, Result};
use zip::read::ZipArchive;
use quick_xml::reader::Reader;
use quick_xml::events::Event;

#[derive(Parser, Debug)]
#[command(name = "xcsv", author, version, about = "Convert XLSX sheets to CSV", long_about = None)]
struct Cli {
    /// Path to the .xlsx file
    #[arg(value_name = "XLSX_PATH")]
    xlsx_path: PathBuf,

    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand, Debug, Clone)]
enum Command {
    /// List sheet names in the workbook
    List,
    /// Export all sheets to CSV files in output directory
    Export {
        /// Output directory (created if missing)
        #[arg(short, long, value_name = "DIR", default_value = ".")]
        out_dir: PathBuf,
    },
}

fn parse_args() -> Cli {
    Cli::parse()
}

fn open_zip(path: &Path) -> Result<ZipArchive<BufReader<File>>> {
    let file = File::open(path).with_context(|| format!("Failed to open {:?}", path))?;
    let reader = BufReader::new(file);
    let zip = ZipArchive::new(reader).context("Failed to read XLSX (zip) archive")?;
    Ok(zip)
}

#[derive(Debug, Clone)]
struct SheetInfo { name: String, path_in_zip: String }

fn tag_eq_ignore_case(actual: &[u8], expect: &str) -> bool {
    actual.eq_ignore_ascii_case(expect.as_bytes()) || actual.ends_with(expect.as_bytes()) || actual.ends_with(expect.to_ascii_lowercase().as_bytes()) || actual.ends_with(expect.to_ascii_uppercase().as_bytes())
}

fn parse_workbook_rels<R: BufRead>(reader: R) -> Result<BTreeMap<String, String>> {
    // Map r:Id -> full path inside zip (xl/worksheets/sheet1.xml)
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);
    let mut buf = Vec::new();
    let mut map = BTreeMap::new();
    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                if tag_eq_ignore_case(e.name().as_ref(), "Relationship") {
                    let mut id = None;
                    let mut target = None;
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"Id" | b"r:Id" => id = Some(String::from_utf8_lossy(&a.value).into_owned()),
                            b"Target" => target = Some(String::from_utf8_lossy(&a.value).into_owned()),
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(target)) = (id, target) {
                        map.insert(id, format!("xl/{}", target.trim_start_matches('/')));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML error in workbook.rels: {}", e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(map)
}

fn parse_workbook<R: BufRead>(reader: R, rels: &BTreeMap<String, String>) -> Result<Vec<SheetInfo>> {
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);
    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                if tag_eq_ignore_case(e.name().as_ref(), "sheet") {
                    let mut name = None;
                    let mut r_id = None;
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"name" => name = Some(String::from_utf8_lossy(&a.value).into_owned()),
                            b"id" | b"r:id" => r_id = Some(String::from_utf8_lossy(&a.value).into_owned()),
                            _ => {}
                        }
                    }
                    if let (Some(name), Some(rid)) = (name, r_id) {
                        if let Some(target) = rels.get(&rid) {
                            sheets.push(SheetInfo { name, path_in_zip: target.clone() });
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML error in workbook.xml: {}", e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(sheets)
}

fn read_shared_strings<R: BufRead>(reader: R) -> Result<Vec<String>> {
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);
    let mut buf = Vec::new();
    let mut strings = Vec::new();
    let mut in_si = false;
    let mut current = String::new();
    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if tag_eq_ignore_case(e.name().as_ref(), "si") {
                    in_si = true;
                    current.clear();
                }
            }
            Ok(Event::End(e)) => {
                if tag_eq_ignore_case(e.name().as_ref(), "si") {
                    strings.push(current.clone());
                    in_si = false;
                }
            }
            Ok(Event::Text(t)) => {
                if in_si { current.push_str(&t.unescape()?); }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML error in sharedStrings: {}", e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(strings)
}

#[derive(Debug, Clone, Copy)]
struct CellRef { col: u32, row: u32 }

fn col_to_index(col: &str) -> u32 {
    let mut n: u32 = 0;
    for b in col.bytes() {
        if !(b'A'..=b'Z').contains(&b) { break; }
        n = n * 26 + ((b - b'A' + 1) as u32);
    }
    n
}

fn parse_cell_ref(s: &str) -> Option<CellRef> {
    let mut col = String::new();
    let mut row = String::new();
    for ch in s.chars() {
        if ch.is_ascii_alphabetic() { col.push(ch.to_ascii_uppercase()); } else { row.push(ch); }
    }
    if col.is_empty() || row.is_empty() { return None; }
    Some(CellRef { col: col_to_index(&col), row: row.parse().ok()? })
}

fn to_lowercase_filename(name: &str) -> String {
    let s: String = name.chars().map(|c| {
        if c.is_ascii_alphanumeric() || c == '-' || c == '_' { c.to_ascii_lowercase() } else { '_' }
    }).collect();
    if s.is_empty() { "sheet".to_string() } else { s }
}

fn export_sheet_xml_to_csv<R: BufRead>(reader: R, shared_strings: &[String], out_path: &Path) -> Result<()> {
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);
    let mut buf = Vec::new();
    let mut wtr = csv::WriterBuilder::new().flexible(true).from_path(out_path)?;

    let mut current_row_idx: u32 = 0;
    let mut row_vals: Vec<String> = Vec::new();
    let mut cell_col: Option<u32> = None;
    let mut cell_type: Option<String> = None;
    let mut cell_val: String = String::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if tag_eq_ignore_case(e.name().as_ref(), "row") {
                    // ensure row index continuity
                    let mut r_attr = None;
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"r" { r_attr = String::from_utf8_lossy(&a.value).parse::<u32>().ok(); }
                    }
                    let next = r_attr.unwrap_or(current_row_idx + 1);
                    while current_row_idx + 1 < next {
                        // write empty row
                        wtr.write_record(std::iter::empty::<String>())?;
                        current_row_idx += 1;
                    }
                    current_row_idx = next;
                    row_vals.clear();
                } else if tag_eq_ignore_case(e.name().as_ref(), "c") {
                    cell_col = None; cell_type = None; cell_val.clear();
                    let mut r_attr = None;
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"r" => r_attr = parse_cell_ref(&String::from_utf8_lossy(&a.value)),
                            b"t" => cell_type = Some(String::from_utf8_lossy(&a.value).into_owned()),
                            _ => {}
                        }
                    }
                    if let Some(cr) = r_attr { cell_col = Some(cr.col); }
                }
            }
            Ok(Event::End(e)) => {
                if tag_eq_ignore_case(e.name().as_ref(), "c") {
                    let col = cell_col.unwrap_or((row_vals.len() as u32) + 1);
                    let needed = col as usize;
                    if row_vals.len() < needed { row_vals.resize(needed, String::new()); }
                    let v = match cell_type.as_deref() {
                        Some("s") => {
                            if let Ok(idx) = cell_val.trim().parse::<usize>() { shared_strings.get(idx).cloned().unwrap_or_default() } else { String::new() }
                        }
                        Some("b") => if cell_val.trim() == "1" { "TRUE".to_string() } else { "FALSE".to_string() },
                        Some("inlineStr") => cell_val.clone(),
                        _ => cell_val.clone(),
                    };
                    row_vals[(col as usize) - 1] = v;
                    cell_col = None; cell_type = None; cell_val.clear();
                } else if tag_eq_ignore_case(e.name().as_ref(), "row") {
                    wtr.write_record(row_vals.iter())?;
                    row_vals.clear();
                }
            }
            Ok(Event::Text(t)) => {
                let txt = t.unescape()?;
                if !txt.is_empty() { cell_val.push_str(&txt); }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML error in worksheet: {}", e)),
            _ => {}
        }
        buf.clear();
    }
    if !row_vals.is_empty() { wtr.write_record(row_vals.iter())?; }
    wtr.flush()?;
    Ok(())
}

fn main() -> Result<()> {
    let cli = parse_args();
    let mut zip = open_zip(&cli.xlsx_path)?;

    match cli.command {
        Command::List => {
            // Stream-parse workbook rels
            let rels_map = {
                let f = zip.by_name("xl/_rels/workbook.xml.rels").context("missing xl/_rels/workbook.xml.rels")?;
                let reader = BufReader::new(f);
                parse_workbook_rels(reader)?
            };
            // Stream-parse workbook
            let sheets = {
                let f = zip.by_name("xl/workbook.xml").context("missing xl/workbook.xml")?;
                let reader = BufReader::new(f);
                parse_workbook(reader, &rels_map)?
            };

            for s in sheets { println!("{}", s.name); }
        }
        Command::Export { out_dir } => {
            std::fs::create_dir_all(&out_dir).context("create output directory")?;

            // Stream-parse shared strings if present
            let shared_strings: Vec<String> = if let Ok(f) = zip.by_name("xl/sharedStrings.xml") {
                let reader = BufReader::new(f);
                read_shared_strings(reader)?
            } else { Vec::new() };

            // Workbook rels and sheets
            let rels_map = {
                let f = zip.by_name("xl/_rels/workbook.xml.rels").context("missing xl/_rels/workbook.xml.rels")?;
                let reader = BufReader::new(f);
                parse_workbook_rels(reader)?
            };
            let sheets = {
                let f = zip.by_name("xl/workbook.xml").context("missing xl/workbook.xml")?;
                let reader = BufReader::new(f);
                parse_workbook(reader, &rels_map)?
            };

            // Export each sheet
            for sheet in sheets {
                let filename = format!("{}.csv", to_lowercase_filename(&sheet.name));
                let out_path = out_dir.join(filename);
                let f = zip.by_name(&sheet.path_in_zip).with_context(|| format!("missing {}", sheet.path_in_zip))?;
                let reader = BufReader::new(f);
                export_sheet_xml_to_csv(reader, &shared_strings, &out_path)?;
                eprintln!("wrote {:?}", out_path);
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn make_minimal_xlsx(files: Vec<(&str, &str)>) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let cursor = std::io::Cursor::new(&mut bytes);
            let mut writer = zip::ZipWriter::new(cursor);
            let options = zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

            // [Content_Types].xml
            writer.start_file("[Content_Types].xml", options).unwrap();
            writer.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#).unwrap();

            // xl/_rels/workbook.xml.rels
            writer.add_directory("xl/_rels/", options).ok();
            writer.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
            writer.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#).unwrap();

            // xl/workbook.xml
            writer.add_directory("xl/", options).ok();
            writer.start_file("xl/workbook.xml", options).unwrap();
            writer.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#).unwrap();

            // xl/worksheets/sheet1.xml
            writer.add_directory("xl/worksheets/", options).ok();
            writer.start_file("xl/worksheets/sheet1.xml", options).unwrap();
            writer.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><v>123</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>1</v></c>
      <c r="C2" t="b"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#).unwrap();

            // xl/sharedStrings.xml
            writer.start_file("xl/sharedStrings.xml", options).unwrap();
            writer.write_all(br#"<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
</sst>"#).unwrap();

            // extra provided files
            for (path, content) in files {
                writer.start_file(path, options).unwrap();
                writer.write_all(content.as_bytes()).unwrap();
            }

            writer.finish().unwrap();
        }
        bytes
    }

    #[test]
    fn list_sheets_from_mock_zip() {
        let data = make_minimal_xlsx(vec![]);
        let cursor = std::io::Cursor::new(data);
        let reader = BufReader::new(cursor);
        let mut zip = ZipArchive::new(reader).expect("zip open");

        let rels = {
            let f = zip.by_name("xl/_rels/workbook.xml.rels").unwrap();
            parse_workbook_rels(BufReader::new(f)).unwrap()
        };
        let sheets = {
            let f = zip.by_name("xl/workbook.xml").unwrap();
            parse_workbook(BufReader::new(f), &rels).unwrap()
        };
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Sheet1");
    }

    #[test]
    fn export_mock_sheet_to_csv() {
        let data = make_minimal_xlsx(vec![]);
        let cursor = std::io::Cursor::new(data);
        let reader = BufReader::new(cursor);
        let mut zip = ZipArchive::new(reader).expect("zip open");

        let shared = {
            let f = zip.by_name("xl/sharedStrings.xml").unwrap();
            read_shared_strings(BufReader::new(f)).unwrap()
        };

        let rels = {
            let f = zip.by_name("xl/_rels/workbook.xml.rels").unwrap();
            parse_workbook_rels(BufReader::new(f)).unwrap()
        };
        let sheets = {
            let f = zip.by_name("xl/workbook.xml").unwrap();
            parse_workbook(BufReader::new(f), &rels).unwrap()
        };

        let dir = tempfile::tempdir().unwrap();
        let out_path = dir.path().join("sheet1.csv");
        let f = zip.by_name(&sheets[0].path_in_zip).unwrap();
        export_sheet_xml_to_csv(BufReader::new(f), &shared, &out_path).unwrap();

        let got = std::fs::read_to_string(out_path).unwrap();
        // Rows: [Hello,123] and [World,"",TRUE] (C2 is bool TRUE, B2 empty)
        let lines: Vec<&str> = got.trim().split('\n').collect();
        assert_eq!(lines[0], "Hello,123");
        assert_eq!(lines[1], "World,,TRUE");
    }
}
