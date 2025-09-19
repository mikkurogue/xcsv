use std::collections::BTreeMap;
use std::fs::File;
use std::io::{BufRead, BufReader};
use std::path::{Path, PathBuf};

use clap::{Parser, Subcommand};
use anyhow::{Context, Result};
use zip::read::ZipArchive;
use quick_xml::reader::Reader;
use quick_xml::events::Event;

// Import functions from lib module
use xcsv::{
    parse_workbook_rels, parse_workbook, read_shared_strings, 
    export_sheet_xml_to_csv, SheetInfo, CellRef, col_to_index, 
    parse_cell_ref, to_lowercase_filename, excel_serial_to_iso_date
};

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
        /// CSV delimiter character
        #[arg(short, long, value_name = "DELIMITER", default_value = ",", value_parser = parse_delimiter)]
        delimiter: u8,
    },
}

fn parse_args() -> Cli {
    Cli::parse()
}

fn parse_delimiter(s: &str) -> Result<u8, String> {
    match s {
        "," => Ok(b','),
        ";" => Ok(b';'),
        _ => Err(format!("Invalid delimiter '{}'. Supported delimiters: ',' (comma) or ';' (semicolon)", s))
    }
}

fn open_zip(path: &Path) -> Result<ZipArchive<BufReader<File>>> {
    let file = File::open(path).with_context(|| format!("Failed to open {:?}", path))?;
    let reader = BufReader::new(file);
    let zip = ZipArchive::new(reader).context("Failed to read XLSX (zip) archive")?;
    Ok(zip)
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
        Command::Export { out_dir, delimiter } => {
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
                export_sheet_xml_to_csv(reader, &shared_strings, &out_path, delimiter)?;
                eprintln!("wrote {:?}", out_path);
            }
        }
    }
    Ok(())
}

