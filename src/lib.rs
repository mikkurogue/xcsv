// Library interface for xcsv - exposes functions for testing

use std::collections::BTreeMap;
use std::io::{BufRead, BufReader};
use std::path::Path;
use anyhow::{Context, Result};
use quick_xml::reader::Reader;
use quick_xml::events::Event;
use chrono;

#[derive(Debug, Clone)]
pub struct SheetInfo { pub name: String, pub path_in_zip: String }

fn tag_eq_ignore_case(actual: &[u8], expect: &str) -> bool {
    actual.eq_ignore_ascii_case(expect.as_bytes()) || actual.ends_with(expect.as_bytes()) || actual.ends_with(expect.to_ascii_lowercase().as_bytes()) || actual.ends_with(expect.to_ascii_uppercase().as_bytes())
}

pub fn parse_workbook_rels<R: BufRead>(reader: R) -> Result<BTreeMap<String, String>> {
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

pub fn parse_workbook<R: BufRead>(reader: R, rels: &BTreeMap<String, String>) -> Result<Vec<SheetInfo>> {
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

pub fn read_shared_strings<R: BufRead>(reader: R) -> Result<Vec<String>> {
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

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CellRef { pub col: u32, pub row: u32 }

pub fn col_to_index(col: &str) -> u32 {
    let mut n: u32 = 0;
    for b in col.bytes() {
        if !(b'A'..=b'Z').contains(&b) { break; }
        n = n * 26 + ((b - b'A' + 1) as u32);
    }
    n
}

pub fn parse_cell_ref(s: &str) -> Option<CellRef> {
    let mut col = String::new();
    let mut row = String::new();
    for ch in s.chars() {
        if ch.is_ascii_alphabetic() { col.push(ch.to_ascii_uppercase()); } else { row.push(ch); }
    }
    if col.is_empty() || row.is_empty() { return None; }
    Some(CellRef { col: col_to_index(&col), row: row.parse().ok()? })
}

pub fn to_lowercase_filename(name: &str) -> String {
    let s: String = name.chars().map(|c| {
        if c.is_ascii_alphanumeric() || c == '-' || c == '_' { c.to_ascii_lowercase() } else { '_' }
    }).collect();
    if s.is_empty() { "sheet".to_string() } else { s }
}

// Excel date/time utilities
// Excel stores dates as serial numbers: days since 1900-01-01 (with 1900 incorrectly treated as leap year)
const EXCEL_EPOCH_DAYS: i32 = 25569; // Days from 1970-01-01 to 1900-01-01
const SECONDS_PER_DAY: f64 = 86400.0;

pub fn excel_serial_to_iso_date(serial: f64) -> Option<String> {
    // Excel serial date: integer part is days since 1900-01-01, fractional part is time
    let days = serial.floor() as i32;
    let time_fraction = serial - days as f64;
    
    // Convert to Unix timestamp (seconds since 1970-01-01)
    let unix_days = days - EXCEL_EPOCH_DAYS;
    let unix_seconds = (unix_days as f64 * SECONDS_PER_DAY) + (time_fraction * SECONDS_PER_DAY);
    
    // Convert to ISO 8601 format
    let timestamp = unix_seconds as i64;
    let datetime = chrono::DateTime::from_timestamp(timestamp, 0)?;
    Some(datetime.format("%Y-%m-%dT%H:%M:%S%.3fZ").to_string())
}

pub fn export_sheet_xml_to_csv<R: BufRead>(reader: R, shared_strings: &[String], out_path: &Path, delimiter: u8) -> Result<()> {
    let mut xml = Reader::from_reader(reader);
    xml.trim_text(true);
    let mut buf = Vec::new();
    let mut wtr = csv::WriterBuilder::new()
        .flexible(true)
        .delimiter(delimiter)
        .from_path(out_path)?;

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
                } else if tag_eq_ignore_case(e.name().as_ref(), "is") {
                    // Inline string container - reset cell_val to collect text
                    cell_val.clear();
                } else if tag_eq_ignore_case(e.name().as_ref(), "t") {
                    // <t> element inside inline string - text will come in Text event
                }
            }
            Ok(Event::End(e)) => {
                if tag_eq_ignore_case(e.name().as_ref(), "c") {
                    let col = cell_col.unwrap_or((row_vals.len() as u32) + 1);
                    let needed = col as usize;
                    if row_vals.len() < needed { row_vals.resize(needed, String::new()); }
                    let v = match cell_type.as_deref() {
                        Some("s") => {
                            // Shared string reference
                            if let Ok(idx) = cell_val.trim().parse::<usize>() { 
                                shared_strings.get(idx).cloned().unwrap_or_default() 
                            } else { 
                                String::new() 
                            }
                        }
                        Some("b") => {
                            // Boolean
                            if cell_val.trim() == "1" { "TRUE".to_string() } else { "FALSE".to_string() }
                        }
                        Some("inlineStr") => {
                            // Inline string (handled in text events)
                            cell_val.clone()
                        }
                        Some("str") => {
                            // Formula result as string
                            cell_val.clone()
                        }
                        Some("e") => {
                            // Error value
                            format!("#ERROR:{}", cell_val)
                        }
                        _ => {
                            // Numeric value - check if it looks like a date
                            if let Ok(num) = cell_val.trim().parse::<f64>() {
                                // Excel dates are typically between 1 (1900-01-01) and ~50000 (2037+)
                                // Be more conservative: only convert if it's in a reasonable date range
                                // and has a reasonable magnitude (not small integers like 123)
                                if num >= 1.0 && num <= 50000.0 && (num >= 1000.0 || num.fract() > 0.0) {
                                    // Could be a date, try to convert
                                    if let Some(iso_date) = excel_serial_to_iso_date(num) {
                                        iso_date
                                    } else {
                                        cell_val.clone()
                                    }
                                } else {
                                    cell_val.clone()
                                }
                            } else {
                                cell_val.clone()
                            }
                        }
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
