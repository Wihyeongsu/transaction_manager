pub mod models;
pub mod send_file;

use models::data::{Data, DataBuilder, Date};
use pdf_extract::extract_text;
use regex::Regex;
use std::error::Error;
use std::path::Path;

pub fn extract_tables_from_pdf(pdf_path: &Path) -> Result<Vec<Data>, Box<dyn Error>> {
    let text = extract_text(pdf_path)?;
    let lines: Vec<&str> = text.lines().collect();

    let mut table: Vec<_> = Vec::new();

    // "1 yyyy.mm.dd hh:mm:ss name 000,000 000,000 000,000 (name) info info (info)"
    let pattern = r"^\s*(\d+)\s+(\d{4}.\d{2}.\d{2} \d{2}:\d{2}:\d{2})\s*(\S+(?:\s+\S+)*)\s+(\d{1,3}(?:,\d{3})*)\s+(\d{1,3}(?:,\d{3})*)\s*(\d{1,3}(?:,\d{3})*)(?:\s*\S+(?:\s+\S+)*)?\s+(\S+)\s+(\S+)(?:\s+(\S+))?\s*$";
    let regex = Regex::new(&pattern)?;
    for line in lines {
        match regex_match(&regex, line) {
            Ok(Some(data)) => table.push(data),
            Ok(None) => {}
            Err(e) => return Err(e.into()),
        }
    }
    Ok(table)
}

pub fn regex_match(regex: &Regex, line: &str) -> Result<Option<Data>, Box<dyn Error>> {
    // println!("{line}");
    if let Some(caps) = regex.captures(line) {
        let mut data_builder = DataBuilder::new();
        let data = data_builder
            .date(Date::new(&caps[2]))
            .cash_in(caps[5].replace(",", "").parse()?)
            .cash_out(caps[4].replace(",", "").parse()?)
            .balance(caps[6].replace(",", "").parse()?)
            .build()?;
        Ok(Some(data))
        // println!("Matched transaction: {:#?}", caps);
    } else {
        Ok(None) // unmatched
    }
}

pub fn cell_name(row: u32, mut col: u32) -> String {
    let mut name = String::new();
    col += 1;
    while col > 0 {
        col -= 1; // 0-based for calculation
        let remainder = (col % 26) as u8;
        name.push((b'A' + remainder) as char);
        col /= 26;
    }
    name = name.chars().rev().collect::<String>();
    format!("{name}{}", row + 1)
}
