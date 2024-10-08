pub mod discord_message;
pub mod format;
pub mod models;
pub mod send_file;
pub mod write_account;
pub mod write_budget;

use format::{format_list, DATE_FORMAT_STR, NUM_FORMAT_STR};
use models::data::{Data, DataBuilder, Date, VariantName};
use regex::Regex;
use rust_xlsxwriter::{ExcelDateTime, FormatBorder, Formula, Worksheet};
use std::collections::HashMap;
use std::error::Error;
use std::fs::File;
use std::io::{BufRead, BufReader};
use std::path::Path;

pub fn extract_tables(file_path: &Path) -> Result<Vec<Data>, Box<dyn Error>> {
    let file = File::open(file_path)?;
    let reader = BufReader::new(file);
    let mut table: Vec<_> = Vec::new();

    // "1 yyyy.mm.dd hh:mm:ss name 000,000 000,000 000,000 (name) info info (info)"
    let pattern = r"^(\d+)\|(\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2})\|([^|]+)\|(\d{1,3}(?:,\d{3})*)\|(\d{1,3}(?:,\d{3})*)\|(\d{1,3}(?:,\d{3})*)\|([^|]*)\|([^|]+)\|([^|]+)\|.*$";
    let regex = Regex::new(&pattern)?;
    for line in reader.lines() {
        match regex_match(&regex, line?.as_str()) {
            Ok(Some(data)) => table.push(data),
            Ok(None) => {}
            Err(e) => return Err(e.into()),
        }
    }
    Ok(table)
}

// 정규표현식 매칭
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

// 월별 데이터 분리
pub fn separate_data(mut table: Vec<Data>) -> Result<Vec<(u8, Vec<Data>)>, Box<dyn Error>> {
    let mut month_data_list: HashMap<u8, Vec<Data>> = HashMap::new();

    let first_month = table.last().expect("first month").date.month;
    let se: (u8, u8) = if first_month == 1 {
        (1, 6)
    } else if first_month == 6 {
        (6, 12)
    } else {
        (0, 0)
    };

    // 데이터가 없는 월 sheet 생성하기 위해
    for month in se.0..=se.1 {
        month_data_list.insert(month, Vec::new());
    }

    while let Some(data) = table.pop() {
        month_data_list
            .entry(data.date.month)
            .or_insert(Vec::new())
            .push(data);
    }

    let mut month_data_list: Vec<(u8, Vec<Data>)> = month_data_list.into_iter().collect();
    month_data_list.sort_by_key(|&(k, _)| k);

    Ok(month_data_list)
}

// 셀 이름 변환
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

    // set sheet cell height
    worksheet
        .set_default_row_height(15)
        .set_row_height(0, 21)?
        .set_row_height(1, 21)?
        .set_row_height(2, 21)?
        .set_row_height(3, 21)?
        .set_row_height(4, 15.5)?
        .set_row_height(5, 15.8)?;

    // set filter
    worksheet.autofilter(4, 0, 4, 7)?;

    // merge cells
    worksheet
        .merge_range(
            0,
            0,
            3,
            1,
            &sheet_name.split(' ').next().unwrap(),
            &format_list(0),
        )?
        .merge_range(0, 3, 0, 5, "금액", &format_list(1))?
        .merge_range(0, 6, 0, 7, "비고", &format_list(1))?
        .merge_range(1, 3, 1, 5, "=", &format_list(6))?
        .merge_range(1, 6, 1, 7, "", &format_list(2))?
        .merge_range(2, 3, 2, 5, "=", &format_list(6))?
        .merge_range(2, 6, 2, 7, "", &format_list(2))?
        .merge_range(3, 3, 3, 5, "=", &format_list(6))?
        .merge_range(3, 6, 3, 7, "", &format_list(3))?
        .merge_range(
            4,
            3,
            4,
            5,
            "금액",
            &format_list(1).clone().set_border_bottom(FormatBorder::None),
        )?;

    //
    worksheet
        .write_with_format(0, 2, "구분", &format_list(1))?
        .write_with_format(1, 2, "수입", &format_list(1))?
        .write_with_format(2, 2, "지출", &format_list(1))?
        .write_with_format(3, 2, "이월금", &format_list(1))?
        .write_with_format(
            4,
            0,
            "날짜",
            &format_list(1).clone().set_border_bottom(FormatBorder::None),
        )?
        .write_with_format(
            4,
            1,
            "사업구분",
            &format_list(1).clone().set_border_bottom(FormatBorder::None),
        )?
        .write_with_format(
            4,
            2,
            "사업명",
            &format_list(1).clone().set_border_bottom(FormatBorder::None),
        )?
        .write_with_format(
            4,
            6,
            "비고",
            &format_list(1).clone().set_border_bottom(FormatBorder::None),
        )?
        .write_with_format(
            4,
            7,
            "영수증번호",
            &format_list(1).clone().set_border_bottom(FormatBorder::None),
        )?
        .write_with_format(
            5,
            0,
            "",
            &format_list(1).clone().set_border_top(FormatBorder::None),
        )?
        .write_with_format(
            5,
            1,
            "",
            &format_list(1).clone().set_border_top(FormatBorder::None),
        )?
        .write_with_format(
            5,
            2,
            "",
            &format_list(1).clone().set_border_top(FormatBorder::None),
        )?
        .write_with_format(5, 3, "수입", &format_list(1))?
        .write_with_format(5, 4, "지출", &format_list(1))?
        .write_with_format(5, 5, "잔고", &format_list(1))?
        .write_with_format(
            5,
            6,
            "",
            &format_list(1).clone().set_border_top(FormatBorder::None),
        )?
        .write_with_format(
            5,
            7,
            "",
            &format_list(1).clone().set_border_top(FormatBorder::None),
        )?;

    Ok(())
}

pub fn write_row_data(
    worksheet: &mut Worksheet,
    i: u32,
    datetime: &ExcelDateTime,
    data: &Data,
) -> Result<(), Box<dyn Error>> {
    // 날짜
    worksheet.write_with_format(
        6 + i,
        0,
        datetime,
        &format_list(2).set_num_format(DATE_FORMAT_STR),
    )?;

    // 사업구분
    worksheet.write_with_format(6 + i, 1, data.business_type.variant_name(), &format_list(2))?;

    // 사업명
    worksheet.write_with_format(
        6 + i,
        2,
        data.business_name.clone().unwrap_or_default(),
        &format_list(2),
    )?;

    // 수입
    worksheet.write_with_format(
        6 + i,
        3,
        data.cash_in,
        &format_list(5).set_num_format(NUM_FORMAT_STR),
    )?;

    // 지출
    worksheet.write_with_format(
        6 + i,
        4,
        data.cash_out,
        &format_list(5).set_num_format(NUM_FORMAT_STR),
    )?;

    // 잔고
    worksheet.write_formula_with_format(
        6 + i,
        5,
        Formula::new(format!(
            "={}+{}-{}",
            match i {
                0 => cell_name(3, 3), // 이월금
                _ => cell_name(5 + i, 5),
            },
            cell_name(6 + i, 3),
            cell_name(6 + i, 4)
        )),
        &format_list(6),
    )?;

    // 비고
    worksheet.write_with_format(
        6 + i,
        6,
        data.remarks.clone().unwrap_or_default(),
        &format_list(2),
    )?;

    // 영수증번호
    worksheet.write_with_format(
        6 + i,
        7,
        data.receipt_num.clone().unwrap_or_default(),
        &format_list(2),
    )?;

    Ok(())
}

pub fn write_data_in_sheet(
    worksheet: &mut Worksheet,
    data_list: &Vec<Data>,
) -> Result<(), Box<dyn Error>> {
    for (i, data) in data_list.iter().enumerate() {
        let datetime = ExcelDateTime::from_ymd(data.date.year, data.date.month, data.date.day)?;
        write_row_data(worksheet, i as u32, &datetime, data)?;
    }

    let len = data_list.len() as u32;
    // 공백 열
    // 날짜
    worksheet.write_with_format(
        6 + len,
        0,
        "",
        &format_list(2).set_num_format(DATE_FORMAT_STR),
    )?;

    // 사업구분
    worksheet.write_with_format(6 + len, 1, "", &format_list(2))?;

    // 사업명
    worksheet.write_with_format(6 + len, 2, "", &format_list(2))?;

    // 수입
    worksheet.write_with_format(
        6 + len,
        3,
        "",
        &format_list(5).set_num_format(NUM_FORMAT_STR),
    )?;

    // 지출
    worksheet.write_with_format(
        6 + len,
        4,
        "",
        &format_list(5).set_num_format(NUM_FORMAT_STR),
    )?;

    // 잔고
    worksheet.write_formula_with_format(
        6 + len,
        5,
        Formula::new(format!(
            "={}+{}-{}",
            match len {
                0 => cell_name(3, 3), // 이월금
                _ => cell_name(5 + len, 5),
            },
            cell_name(6 + len, 3),
            cell_name(6 + len, 4)
        )),
        &format_list(6),
    )?;

    // 비고
    worksheet.write_with_format(6 + len, 6, "", &format_list(2))?;

    // 영수증번호
    worksheet.write_with_format(6 + len, 7, "", &format_list(2))?;

    // 계
    worksheet
        .merge_range(7 + len, 0, 7 + len, 2, "계", &format_list(4))?
        .write_formula_with_format(
            7 + len,
            3,
            Formula::new(format!(
                "=SUM({}:{})",
                cell_name(6, 3),
                cell_name(6 + len, 3)
            )),
            &format_list(7).set_num_format(NUM_FORMAT_STR),
        )?
        .write_formula_with_format(
            7 + len,
            4,
            Formula::new(format!(
                "=SUM({}:{})",
                cell_name(6, 4),
                cell_name(6 + len, 4)
            )),
            &format_list(7).set_num_format(NUM_FORMAT_STR),
        )?
        .write_formula_with_format(
            7 + len,
            5,
            Formula::new(format!("={}", cell_name(6 + len, 5),)),
            &format_list(7).set_num_format(NUM_FORMAT_STR),
        )?
        .write_with_format(7 + len, 6, "", &format_list(4))?
        .write_with_format(7 + len, 7, "", &format_list(4))?;

    Ok(())
}
