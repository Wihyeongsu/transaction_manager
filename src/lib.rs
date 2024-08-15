use pdf_extract::extract_text;
use regex::Regex;
use std::error::Error;
use std::path::Path;

#[derive(Debug, Clone)]
pub struct Transaction {
    pub index: String,
    pub date_time: String,
    pub name1: String,
    pub amount1: String,
    pub amount2: String,
    pub amount3: String,
    pub name2: Option<String>,
    pub transaction_type: String,
    pub at: String,
    pub distinct: Option<String>,
}

#[derive(Debug, PartialEq)]
pub struct Date {
    pub year: u16,
    pub month: u8,
    pub day: u8,
}

impl Date {
    pub fn new(date: &str) -> Date {
        let caps = Regex::new(r"(\d{4})\.(\d{2})\.(\d{2})")
            .unwrap()
            .captures(date)
            .unwrap();

        Date {
            year: caps[1].parse().unwrap(),
            month: caps[2].parse().unwrap(),
            day: caps[3].parse().unwrap(),
        }
    }
}

#[derive(Debug, PartialEq)]
pub struct FetchedData {
    pub date: Date,
    pub cash_in: u32,
    pub cash_out: u32,
    pub balance: u32,
}

pub fn extract_tables_from_pdf(pdf_path: &Path) -> Result<Vec<FetchedData>, Box<dyn Error>> {
    let text = extract_text(pdf_path)?;
    let lines: Vec<&str> = text.lines().collect();

    let mut table: Vec<_> = Vec::new();

    // "1 yyyy.mm.dd hh:mm:ss name 000,000 000,000 000,000 (name) info info (info)"
    let pattern = r"^\s*(\d+)\s+(\d{4}.\d{2}.\d{2} \d{2}:\d{2}:\d{2})\s*(\S+(?:\s+\S+)*)\s+(\d{1,3}(?:,\d{3})*)\s+(\d{1,3}(?:,\d{3})*)\s*(\d{1,3}(?:,\d{3})*)(?:\s*\S+(?:\s+\S+)*)?\s+(\S+)\s+(\S+)(?:\s+(\S+))?\s*$";
    let regex = Regex::new(&pattern)?;
    for line in lines {
        match regex_match(&regex, line) {
            Ok(Some(fetched_data)) => table.push(fetched_data),
            Ok(None) => {}
            Err(e) => return Err(e.into()),
        }
    }
    Ok(table)
}

pub fn regex_match(regex: &Regex, line: &str) -> Result<Option<FetchedData>, Box<dyn Error>> {
    // println!("{line}");
    if let Some(caps) = regex.captures(line) {
        // let transaction = Transaction {
        //     index: caps[1].parse().unwrap(),
        //     date_time: caps[2].to_string(),
        //     name1: caps[3].to_string(),
        //     amount1: caps[4].to_string(),
        //     amount2: caps[5].to_string(),
        //     amount3: caps[6].to_string(),
        //     name2: caps.get(7).map(|m| m.as_str().to_string()),
        //     transaction_type: caps[8].to_string(),
        //     at: caps[9].to_string(),
        //     distinct: caps.get(10).map(|m| m.as_str().to_string()),
        // };
        Ok(Some(FetchedData {
            date: Date::new(&caps[2]),
            cash_in: caps[5].replace(",", "").parse()?,
            cash_out: caps[4].replace(",", "").parse()?,
            balance: caps[6].replace(",", "").parse()?,
        }))
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
