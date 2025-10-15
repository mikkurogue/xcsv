use std::io::BufReader;
use std::path::PathBuf;

use anyhow::{Context, Result};
use clap::{Parser, Subcommand};
use libxcsv::{
    StyleInfo, export_sheet_xml_to_csv, open_zip, parse_styles, parse_workbook,
    parse_workbook_rels, read_shared_strings, to_lowercase_filename,
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
        _ => Err(format!(
            "Invalid delimiter '{}'. Supported delimiters: ',' (comma) or ';' (semicolon)",
            s
        )),
    }
}

fn main() -> Result<()> {
    let cli = parse_args();
    let mut zip = open_zip(&cli.xlsx_path)?;

    match cli.command {
        Command::List => {
            // Stream-parse workbook rels
            let rels_map = {
                let f = zip
                    .by_name("xl/_rels/workbook.xml.rels")
                    .context("missing xl/_rels/workbook.xml.rels")?;
                let reader = BufReader::new(f);
                parse_workbook_rels(reader)?
            };
            // Stream-parse workbook
            let (sheets, _) = {
                let f = zip
                    .by_name("xl/workbook.xml")
                    .context("missing xl/workbook.xml")?;
                let reader = BufReader::new(f);
                parse_workbook(reader, &rels_map)?
            };

            for s in sheets {
                println!("{}", s.name);
            }
        }
        Command::Export { out_dir, delimiter } => {
            std::fs::create_dir_all(&out_dir).context("create output directory")?;

            // Stream-parse shared strings if present
            let shared_strings: Vec<String> = if let Ok(f) = zip.by_name("xl/sharedStrings.xml") {
                let reader = BufReader::new(f);
                read_shared_strings(reader)?
            } else {
                Vec::new()
            };

            // Stream-parse styles if present
            let styles: Vec<StyleInfo> = if let Ok(f) = zip.by_name("xl/styles.xml") {
                let reader = BufReader::new(f);
                parse_styles(reader)?
            } else {
                Vec::new()
            };

            // Workbook rels and sheets
            let rels_map = {
                let f = zip
                    .by_name("xl/_rels/workbook.xml.rels")
                    .context("missing xl/_rels/workbook.xml.rels")?;
                let reader = BufReader::new(f);
                parse_workbook_rels(reader)?
            };
            let (sheets, is_1904) = {
                let f = zip
                    .by_name("xl/workbook.xml")
                    .context("missing xl/workbook.xml")?;
                let reader = BufReader::new(f);
                parse_workbook(reader, &rels_map)?
            };

            // Export each sheet
            for sheet in sheets {
                let filename = format!("{}.csv", to_lowercase_filename(&sheet.name));
                let out_path = out_dir.join(filename);
                let f = zip
                    .by_name(&sheet.path_in_zip)
                    .with_context(|| format!("missing {}", sheet.path_in_zip))?;
                let reader = BufReader::new(f);
                export_sheet_xml_to_csv(
                    reader,
                    &shared_strings,
                    &styles,
                    is_1904,
                    &out_path,
                    delimiter,
                )?;
                eprintln!("wrote {:?}", out_path);
            }
        }
    }
    Ok(())
}
