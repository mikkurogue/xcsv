use std::fs::File;
use std::io::BufReader;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result};
use clap::{Parser, Subcommand};
use tracing::{debug, info, instrument, warn};
use tracing_subscriber::{fmt, prelude::*, EnvFilter};
use zip::read::ZipArchive;

// Import functions from lib module
use xcsv::{
    export_sheet_xml_to_csv, parse_styles, parse_workbook, parse_workbook_rels,
    read_shared_strings, to_lowercase_filename, StyleInfo,
};

#[derive(Parser, Debug)]
#[command(name = "xcsv", author, version, about = "Convert XLSX sheets to CSV", long_about = None)]
struct Cli {
    /// Path to the .xlsx file
    #[arg(value_name = "XLSX_PATH")]
    xlsx_path: PathBuf,

    /// Enable debug tracing output
    #[arg(long, global = true)]
    trace: bool,

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

#[instrument(level = "debug", skip_all)]
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

#[instrument(level = "debug", skip(path), fields(path = %path.display()))]
fn open_zip(path: &Path) -> Result<ZipArchive<BufReader<File>>> {
    debug!("Opening XLSX file");
    let file = File::open(path).with_context(|| format!("Failed to open {:?}", path))?;
    let reader = BufReader::new(file);
    let zip = ZipArchive::new(reader).context("Failed to read XLSX (zip) archive")?;
    debug!(file_count = zip.len(), "Opened XLSX archive");
    Ok(zip)
}

fn main() -> Result<()> {
    let cli = parse_args();

    // Initialize tracing based on --trace flag
    // Logs go to stderr to not interfere with stdout output
    let filter = if cli.trace {
        EnvFilter::new("xcsv=debug")
    } else {
        EnvFilter::new("xcsv=warn")
    };

    tracing_subscriber::registry()
        .with(fmt::layer().with_writer(std::io::stderr))
        .with(filter)
        .init();

    info!(xlsx_path = %cli.xlsx_path.display(), "Starting xcsv");
    let mut zip = open_zip(&cli.xlsx_path)?;

    match cli.command {
        Command::List => {
            info!("Listing sheets");
            // Stream-parse workbook rels
            let rels_map = {
                debug!("Parsing workbook relationships");
                let f = zip
                    .by_name("xl/_rels/workbook.xml.rels")
                    .context("missing xl/_rels/workbook.xml.rels")?;
                let reader = BufReader::new(f);
                parse_workbook_rels(reader)?
            };
            // Stream-parse workbook
            let (sheets, _) = {
                debug!("Parsing workbook");
                let f = zip
                    .by_name("xl/workbook.xml")
                    .context("missing xl/workbook.xml")?;
                let reader = BufReader::new(f);
                parse_workbook(reader, &rels_map)?
            };

            info!(sheet_count = sheets.len(), "Found sheets");
            for s in sheets {
                println!("{}", s.name);
            }
        }
        Command::Export { out_dir, delimiter } => {
            info!(out_dir = %out_dir.display(), delimiter = ?char::from(delimiter), "Exporting sheets");
            std::fs::create_dir_all(&out_dir).context("create output directory")?;

            // Stream-parse shared strings if present
            let shared_strings: Vec<String> = if let Ok(f) = zip.by_name("xl/sharedStrings.xml") {
                debug!("Parsing shared strings");
                let reader = BufReader::new(f);
                let strings = read_shared_strings(reader)?;
                debug!(count = strings.len(), "Loaded shared strings");
                strings
            } else {
                warn!("No shared strings found");
                Vec::new()
            };

            // Stream-parse styles if present
            let styles: Vec<StyleInfo> = if let Ok(f) = zip.by_name("xl/styles.xml") {
                debug!("Parsing styles");
                let reader = BufReader::new(f);
                let s = parse_styles(reader)?;
                debug!(count = s.len(), "Loaded styles");
                s
            } else {
                warn!("No styles found");
                Vec::new()
            };

            // Workbook rels and sheets
            let rels_map = {
                debug!("Parsing workbook relationships");
                let f = zip
                    .by_name("xl/_rels/workbook.xml.rels")
                    .context("missing xl/_rels/workbook.xml.rels")?;
                let reader = BufReader::new(f);
                parse_workbook_rels(reader)?
            };
            let (sheets, is_1904) = {
                debug!("Parsing workbook");
                let f = zip
                    .by_name("xl/workbook.xml")
                    .context("missing xl/workbook.xml")?;
                let reader = BufReader::new(f);
                parse_workbook(reader, &rels_map)?
            };

            info!(
                sheet_count = sheets.len(),
                is_1904 = is_1904,
                "Found sheets"
            );
            // Export each sheet
            for sheet in &sheets {
                debug!(sheet_name = %sheet.name, path = %sheet.path_in_zip, "Exporting sheet");
                let filename = format!("{}.csv", to_lowercase_filename(&sheet.name));
                let out_path = out_dir.join(&filename);
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
                info!(output = %out_path.display(), "Wrote CSV");
            }
        }
    }
    Ok(())
}
