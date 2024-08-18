pub mod discord_message;
pub mod models;
pub mod send_file;

use models::data::{Data, DataBuilder, Date};
use pdf_extract::extract_text;
use regex::Regex;
use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, Worksheet};
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

// 월별 정산서 템플릿
pub fn sheet_template(worksheet: &mut Worksheet, sheet_name: &str) -> Result<(), Box<dyn Error>> {
    // sheet title
    worksheet.set_name(sheet_name)?;

    // set sheet cell width
    worksheet
        .set_column_width(0, 8.64)?
        .set_column_width(1, 11.91)?
        .set_column_width(2, 13.64)?
        .set_column_width(3, 12)?
        .set_column_width(4, 12)?
        .set_column_width(5, 13)?
        .set_column_width(6, 54.91)?
        .set_column_width(7, 17.36)?;

    let bg_lavender = Color::RGB(0xCCC1DE); // lavender
    let bg_gray = Color::RGB(0xD8D8D8); // gray

    // month
    let format1 = Format::new()
        .set_align(FormatAlign::VerticalCenter)
        .set_background_color(bg_lavender)
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_font_name("Batangche")
        .set_font_size(15)
        .set_bold();

    // schema
    let format2 = Format::new()
        .set_align(FormatAlign::VerticalCenter)
        .set_align(FormatAlign::Center)
        .set_background_color(bg_lavender)
        .set_border(FormatBorder::Thin)
        .set_font_name("Batangche")
        .set_font_size(10)
        .set_bold();

    // text
    let format3 = Format::new()
        .set_align(FormatAlign::VerticalCenter)
        .set_align(FormatAlign::Center)
        .set_border(FormatBorder::Thin)
        .set_font_name("Batangche")
        .set_font_size(10);

    // gray text
    let format4 = Format::new()
        .set_align(FormatAlign::VerticalCenter)
        .set_align(FormatAlign::Center)
        .set_background_color(bg_gray)
        .set_border(FormatBorder::Thin)
        .set_font_name("Batangche")
        .set_font_size(10);

    // lavender text
    let format5 = Format::new()
        .set_align(FormatAlign::VerticalCenter)
        .set_align(FormatAlign::Center)
        .set_background_color(bg_lavender)
        .set_border(FormatBorder::Thin)
        .set_font_name("Batangche")
        .set_font_size(10);

    // num
    let format6 = Format::new()
        .set_align(FormatAlign::VerticalCenter)
        .set_align(FormatAlign::Center)
        .set_border(FormatBorder::Thin)
        .set_font_name("Batangche")
        .set_font_size(10)
        .set_num_format("_-₩* #,##0_-;-₩* #,##0_-;_-₩* \" - \"_-;_-@");

    // gray formula
    let format7 = Format::new()
        .set_align(FormatAlign::VerticalCenter)
        .set_align(FormatAlign::Center)
        .set_background_color(bg_gray)
        .set_border(FormatBorder::Thin)
        .set_font_name("Batangche")
        .set_font_size(10)
        .set_num_format("_-₩* #,##0_-;-₩* #,##0_-;_-₩* \" - \"_-;_-@");

    // lavender formula
    let format8 = Format::new()
        .set_align(FormatAlign::VerticalCenter)
        .set_align(FormatAlign::Center)
        .set_background_color(bg_lavender)
        .set_border(FormatBorder::Thin)
        .set_font_name("Batangche")
        .set_font_size(10)
        .set_num_format("_-₩* #,##0_-;-₩* #,##0_-;_-₩* \" - \"_-;_-@");

    // merge cells
    worksheet
        .merge_range(0, 0, 3, 1, &sheet_name.split(' ').next().unwrap(), &format1)?
        .merge_range(0, 3, 0, 5, "금액", &format2)?
        .merge_range(0, 6, 0, 7, "비고", &format2)?
        .merge_range(1, 3, 1, 5, "=", &format7)?
        .merge_range(1, 6, 1, 7, "", &format3)?
        .merge_range(2, 3, 2, 5, "=", &format7)?
        .merge_range(2, 6, 2, 7, "", &format3)?
        .merge_range(3, 3, 3, 5, "=", &format7)?
        .merge_range(3, 6, 3, 7, "", &format4)?
        .merge_range(4, 3, 4, 5, "금액", &format2)?;

    //
    worksheet
        .write_with_format(0, 2, "구분", &format2)?
        .write_with_format(1, 2, "수입", &format2)?
        .write_with_format(2, 2, "지출", &format2)?
        .write_with_format(3, 2, "이월금", &format2)?
        .write_with_format(4, 0, "날짜", &format2)?
        .write_with_format(4, 1, "사업구분", &format2)?
        .write_with_format(4, 2, "사업명", &format2)?
        .write_with_format(4, 6, "비고", &format2)?
        .write_with_format(4, 7, "영수증번호", &format2)?
        .write_with_format(5, 0, "", &format2)?
        .write_with_format(5, 1, "", &format2)?
        .write_with_format(5, 2, "", &format2)?
        .write_with_format(5, 3, "수입", &format2)?
        .write_with_format(5, 4, "지출", &format2)?
        .write_with_format(5, 5, "잔고", &format2)?
        .write_with_format(5, 6, "", &format2)?
        .write_with_format(5, 7, "", &format2)?;

    Ok(())
}
