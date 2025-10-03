// Library interface for xcsv - exposes functions for testing

use anyhow::Result;
use chrono;

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::collections::BTreeMap;
use std::io::BufRead;
use std::path::Path;

#[derive(Debug, Clone)]
pub struct SheetInfo {
    pub name: String,
    pub path_in_zip: String,
}

#[derive(Debug, Clone, Default)]
pub struct StyleInfo {
    pub is_date: bool,
}

pub fn parse_styles<R: BufRead>(reader: R) -> Result<Vec<StyleInfo>> {
    let mut xml = Reader::from_reader(reader);
    let mut buf = Vec::new();
    let mut styles = Vec::new();
    let mut num_fmts = BTreeMap::new();
    let mut in_cell_xfs = false;

    // Helper closure to process attributes of an <xf> tag
    let process_xf = |attrs: quick_xml::events::attributes::Attributes,
                      num_fmts: &BTreeMap<u32, String>|
     -> Result<StyleInfo> {
        let mut style = StyleInfo::default();
        let mut num_fmt_id_attr = None;
        let mut apply_num_fmt = true;

        for a in attrs.flatten() {
            match a.key.as_ref() {
                b"numFmtId" => {
                    num_fmt_id_attr = String::from_utf8_lossy(&a.value).parse::<u32>().ok();
                }
                b"applyNumberFormat" => {
                    apply_num_fmt =
                        String::from_utf8_lossy(&a.value).parse::<u32>().ok() == Some(1);
                }
                _ => {}
            }
        }

        if apply_num_fmt {
            if let Some(id) = num_fmt_id_attr {
                // Check built-in formats
                let is_builtin_date =
                    matches!(id, 14..=22 | 27..=36 | 45..=47 | 50..=58 | 67..=71 | 75..=81);
                if is_builtin_date {
                    style.is_date = true;
                } else if let Some(format_code) = num_fmts.get(&id) {
                    // Check custom formats
                    let lower = format_code.to_lowercase();
                    if (lower.contains('y') || lower.contains('d') || lower.contains('m'))
                        && !lower.contains('#')
                    {
                        style.is_date = true;
                    }
                }
            }
        }
        Ok(style)
    };

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"numFmt" => {
                    let mut num_fmt_id = None;
                    let mut format_code = None;
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"numFmtId" => {
                                num_fmt_id =
                                    String::from_utf8_lossy(&a.value).parse::<u32>().ok();
                            }
                            b"formatCode" => {
                                format_code =
                                    Some(String::from_utf8_lossy(&a.value).into_owned());
                            }
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(code)) = (num_fmt_id, format_code) {
                        num_fmts.insert(id, code);
                    }
                    xml.read_to_end_into(e.name(), &mut Vec::new())?;
                }
                b"cellXfs" => in_cell_xfs = true,
                b"xf" if in_cell_xfs => {
                    styles.push(process_xf(e.attributes(), &num_fmts)?);
                    xml.read_to_end_into(e.name(), &mut Vec::new())?;
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"numFmt" => {
                    let mut num_fmt_id = None;
                    let mut format_code = None;
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"numFmtId" => {
                                num_fmt_id =
                                    String::from_utf8_lossy(&a.value).parse::<u32>().ok();
                            }
                            b"formatCode" => {
                                format_code =
                                    Some(String::from_utf8_lossy(&a.value).into_owned());
                            }
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(code)) = (num_fmt_id, format_code) {
                        num_fmts.insert(id, code);
                    }
                }
                b"xf" if in_cell_xfs => {
                    styles.push(process_xf(e.attributes(), &num_fmts)?);
                }
                _ => {}
            },
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"cellXfs" {
                    in_cell_xfs = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML error in styles: {}", e)),
            _ => {}
        }
        buf.clear();
    }
    Ok(styles)
}

fn tag_eq_ignore_case(actual: &[u8], expect: &str) -> bool {
    actual.eq_ignore_ascii_case(expect.as_bytes())
        || actual.ends_with(expect.as_bytes())
        || actual.ends_with(expect.to_ascii_lowercase().as_bytes())
        || actual.ends_with(expect.to_ascii_uppercase().as_bytes())
}

pub fn parse_workbook_rels<R: BufRead>(reader: R) -> Result<BTreeMap<String, String>> {
    // Map r:Id -> full path inside zip (xl/worksheets/sheet1.xml)
    let mut xml = Reader::from_reader(reader);
    // xml.config_mut().trim_text(true);
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
                            b"Id" | b"r:Id" => {
                                id = Some(String::from_utf8_lossy(&a.value).into_owned())
                            }
                            b"Target" => {
                                target = Some(String::from_utf8_lossy(&a.value).into_owned())
                            }
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

pub fn parse_workbook<R: BufRead>(
    reader: R,
    rels: &BTreeMap<String, String>,
) -> Result<(Vec<SheetInfo>, bool)> {
    let mut xml = Reader::from_reader(reader);
    // xml.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    let mut is_1904 = false;
    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => match e.name().as_ref() {
                b"sheet" => {
                    let mut name = None;
                    let mut r_id = None;
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"name" => name = Some(String::from_utf8_lossy(&a.value).into_owned()),
                            b"id" | b"r:id" => {
                                r_id = Some(String::from_utf8_lossy(&a.value).into_owned())
                            }
                            _ => {}
                        }
                    }
                    if let (Some(name), Some(rid)) = (name, r_id) {
                        if let Some(target) = rels.get(&rid) {
                            sheets.push(SheetInfo {
                                name,
                                path_in_zip: target.clone(),
                            });
                        }
                    }
                }
                b"workbookPr" => {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"date1904" {
                            if let Ok(val) = a.decode_and_unescape_value(&xml) {
                                is_1904 = val == "1" || val == "true";
                            }
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML error in workbook.xml: {}", e)),
            _ => {}
        }
        buf.clear();
    }
    Ok((sheets, is_1904))
}

pub fn read_shared_strings<R: BufRead>(reader: R) -> Result<Vec<String>> {
    let mut xml = Reader::from_reader(reader);
    // xml.config_mut().trim_text(true);
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
                if in_si {
                    // Due to quick-xml 0.38.3 (i assume 0.37+)
                    // The config is unescaping everything way too early.
                    // So we have reverted to 0.31.0 to have a functioning parser
                    // to show correct characters like angle brackets.
                    current.push_str(&t.unescape()?);
                }
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
pub struct CellRef {
    pub col: u32,
    pub row: u32,
}

pub fn col_to_index(col: &str) -> u32 {
    let mut n: u32 = 0;
    for b in col.bytes() {
        if !(b'A'..=b'Z').contains(&b) {
            break;
        }
        n = n * 26 + ((b - b'A' + 1) as u32);
    }
    n
}

pub fn parse_cell_ref(s: &str) -> Option<CellRef> {
    let mut col = String::new();
    let mut row = String::new();
    for ch in s.chars() {
        if ch.is_ascii_alphabetic() {
            col.push(ch.to_ascii_uppercase());
        } else {
            row.push(ch);
        }
    }
    if col.is_empty() || row.is_empty() {
        return None;
    }
    Some(CellRef {
        col: col_to_index(&col),
        row: row.parse().ok()?,
    })
}

pub fn to_lowercase_filename(name: &str) -> String {
    let s: String = name
        .chars()
        .map(|c| {
            if c.is_ascii_alphanumeric() || c == '-' || c == '_' {
                c.to_ascii_lowercase()
            } else {
                '_'
            }
        })
        .collect();
    if s.is_empty() { "sheet".to_string() } else { s }
}

// Excel date/time utilities
// Excel stores dates as serial numbers: days since 1900-01-01 (with 1900 incorrectly treated as leap year)

const SECONDS_PER_DAY: f64 = 86400.0;

pub fn excel_serial_to_iso_date(serial: f64, is_1904: bool) -> Option<String> {
    let excel_epoch_days = if is_1904 {
        24107 // Days from 1970-01-01 to 1904-01-01
    } else {
        25569 // Days from 1970-01-01 to 1900-01-01
    };

    let days = serial.floor() as i32;
    let time_fraction = serial - days as f64;

    // In the 1900 system, Excel incorrectly treats 1900 as a leap year.
    // This means any date after Feb 28, 1900, is off by one.
    // Serial numbers 1-59 are unaffected. 60 is "Feb 29, 1900". 61 is Mar 1, 1900.
    // Our epoch calculation for 1900 already accounts for this by starting from 1899-12-30,
    // so we don't need a special adjustment here if using a correct epoch day count.
    // The constant 25569 = days between 1970-01-01 and 1899-12-30.

    let unix_days = days - excel_epoch_days;
    let unix_seconds =
        (unix_days as f64 * SECONDS_PER_DAY) + (time_fraction * SECONDS_PER_DAY).round();

    let datetime = chrono::DateTime::from_timestamp(unix_seconds as i64, 0)?;
    Some(datetime.format("%Y-%m-%dT%H:%M:%S%.3fZ").to_string())
}

pub fn export_sheet_xml_to_csv<R: BufRead>(
    reader: R,
    shared_strings: &[String],
    styles: &[StyleInfo],
    is_1904: bool,
    out_path: &Path,
    delimiter: u8,
) -> Result<()> {
    let mut xml = Reader::from_reader(reader);
    let mut buf = Vec::new();
    let mut wtr = csv::WriterBuilder::new()
        .flexible(true)
        .delimiter(delimiter)
        .from_path(out_path)?;

    let mut num_columns: Option<usize> = None;
    let mut current_row_idx: u32 = 0;
    let mut row_vals: Vec<String> = Vec::new();
    let mut cell_col: Option<u32> = None;
    let mut cell_type: Option<String> = None;
    let mut cell_style_idx: Option<u32> = None;
    let mut cell_val: String = String::new();

    loop {
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if tag_eq_ignore_case(e.name().as_ref(), "row") {
                    let mut r_attr = None;
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"r" {
                            r_attr = String::from_utf8_lossy(&a.value).parse::<u32>().ok();
                        }
                    }
                    let next = r_attr.unwrap_or(current_row_idx + 1);
                    while current_row_idx + 1 < next {
                        wtr.write_record(std::iter::empty::<String>())?;
                        current_row_idx += 1;
                    }
                    current_row_idx = next;
                    row_vals.clear();
                } else if tag_eq_ignore_case(e.name().as_ref(), "c") {
                    cell_col = None;
                    cell_type = None;
                    cell_val.clear();
                    cell_style_idx = None;
                    let mut r_attr: Option<CellRef> = None;
                    for a in e.attributes().flatten() {
                        match a.key.as_ref() {
                            b"r" => {
                                r_attr = parse_cell_ref(&String::from_utf8_lossy(&a.value));
                            }
                            b"t" => {
                                cell_type = Some(String::from_utf8_lossy(&a.value).into_owned())
                            }
                            b"s" => {
                                cell_style_idx =
                                    String::from_utf8_lossy(&a.value).parse::<u32>().ok();
                            }
                            _ => {}
                        }
                    }
                    if let Some(cr) = r_attr {
                        cell_col = Some(cr.col);
                    }
                } else if tag_eq_ignore_case(e.name().as_ref(), "is") {
                    cell_val.clear();
                } else if tag_eq_ignore_case(e.name().as_ref(), "t") {
                    // text will come in Text event
                }
            }
            Ok(Event::End(e)) => {
                if tag_eq_ignore_case(e.name().as_ref(), "c") {
                    let col = cell_col.unwrap_or((row_vals.len() as u32) + 1);
                    let needed = col as usize;
                    if row_vals.len() < needed {
                        row_vals.resize(needed, String::new());
                    }

                    let v = match cell_type.as_deref() {
                        Some("s") => {
                            if let Ok(idx) = cell_val.trim().parse::<usize>() {
                                shared_strings.get(idx).cloned().unwrap_or_default()
                            } else {
                                String::new()
                            }
                        }
                        Some("b") => if cell_val.trim() == "1" {
                            "TRUE"
                        } else {
                            "FALSE"
                        }
                        .to_string(),
                        Some("inlineStr") | Some("str") => cell_val.clone(),
                        Some("e") => {
                            format!("#ERROR:{}", cell_val)
                        }
                        _ => {
                            // Numeric value
                            match cell_val.trim().parse::<f64>() {
                                Ok(num) => {
                                    let is_date_style = cell_style_idx
                                        .and_then(|idx| styles.get(idx as usize))
                                        .is_some_and(|style_info| style_info.is_date);

                                    if is_date_style {
                                        excel_serial_to_iso_date(num, is_1904).unwrap_or_else(|| cell_val.clone())
                                    } else {
                                        cell_val.clone()
                                    }
                                }
                                Err(_) => cell_val.clone(),
                            }
                        }
                    };
                    row_vals[(col as usize) - 1] = v;

                    cell_col = None;
                    cell_type = None;
                    cell_val.clear();
                    cell_style_idx = None;
                } else if tag_eq_ignore_case(e.name().as_ref(), "row") {
                    if num_columns.is_none() {
                        let last_non_empty = row_vals.iter().rposition(|c| !c.is_empty());
                        num_columns = Some(last_non_empty.map_or(0, |i| i + 1));
                    }
                    if let Some(n) = num_columns {
                        if row_vals.len() < n {
                            row_vals.resize(n, String::new());
                        }
                    }
                    wtr.write_record(row_vals.iter())?;
                    row_vals.clear();
                }
            }
            Ok(Event::Text(t)) => {
                let txt = t.unescape()?;
                if !txt.is_empty() {
                    cell_val.push_str(&txt);
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(anyhow::anyhow!("XML error in worksheet: {}", e)),
            _ => {}
        }
        buf.clear();
    }
    if !row_vals.is_empty() {
        wtr.write_record(row_vals.iter())?;
    }
    wtr.flush()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use std::io::BufReader;
    use tempfile::NamedTempFile;

    #[test]
    fn test_geo_coordinate_parsing_from_xml() {
        let xml_data = r#"
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
        </worksheet>
        "#;
        let shared_strings = vec![
            "origin_latitude".to_string(),
            "origin_longitude".to_string(),
        ];
        let reader = BufReader::new(xml_data.as_bytes());
        let temp_file = NamedTempFile::new().unwrap();
        let out_path = temp_file.path();

        export_sheet_xml_to_csv(reader, &shared_strings, &[], false, out_path, b',').unwrap();

        let csv_content = fs::read_to_string(out_path).unwrap();
        let expected_content = "origin_latitude,origin_longitude\n10.123,-20.456\n";
        assert_eq!(csv_content, expected_content);
    }
}