use rust_xlsxwriter::*;
use std::cell;
use std::error::Error;
use std::path::Path;
use transaction_manager::discord_message::{fetch_data, send_discord_xlsx};
use transaction_manager::format::format_list;
use transaction_manager::models::data;
use transaction_manager::send_file::run_shell_command;
use transaction_manager::{
    cell_name, extract_tables_from_pdf, separate_data, sheet_template, write_data_in_sheet,
};

// 월별 transaction 분류
// 병렬로 sheet 작성
// 이후 workbook에 sheet 추가

#[tokio::main]
async fn main() -> Result<(), Box<dyn Error>> {
    let pdf_path = Path::new("example/account2.pdf");
    let mut table = extract_tables_from_pdf(pdf_path)?;

    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    let month_data_list = separate_data(table);
    let mut data_size = Vec::with_capacity(month_data_list.len()); // 작월 이월금을 가져오기 위해 사용

    for (month, data_list) in month_data_list.iter() {
        let sheet_name = month.to_string() + "월 정산서";
        let mut worksheet = Worksheet::new();
        sheet_template(&mut worksheet, sheet_name.as_str())?;

        // write data
        write_data_in_sheet(&mut worksheet, data_list)?;

        // write schema formula
        worksheet
            // 수입
            .write_formula_with_format(
                1,
                3,
                Formula::new(format!("={}", cell_name(6 + data_list.len() as u32, 3))),
                &format_list(6),
            )?
            // 지출
            .write_formula_with_format(
                2,
                3,
                Formula::new(format!("={}", cell_name(6 + data_list.len() as u32, 4))),
                &format_list(6),
            )?;
        // 이월금
        if workbook.worksheets().is_empty() {
            worksheet
                .write_formula_with_format(
                    3,
                    3,
                    Formula::new(format!(
                        "={}년도 제{}회기 예산안'!B7",
                        data_list[0].date.year,
                        match month {
                            // 첫 sheet에서 검사하는 단계라 overlap 문제는 상관없을 것 같긴 한데
                            // 혹시 모르기는 하니까 이 부분은 가이드라인 감사 대상 기간을 제대로 확인해서 방식을 바꾸든가 해야지
                            6..=12 => 2,
                            1..=6 => 1,
                            _ => 0,
                        }
                    )),
                    &format_list(6),
                )?
                .write_with_format(3, 6, "전단위 인수인계 금액", &format_list(3))?;
        } else {
            worksheet
                .write_formula_with_format(
                    3,
                    3,
                    Formula::new(format!(
                        "='{}월 정산서'!{}",
                        month - 1,
                        cell_name(6 + data_size.last().unwrap(), 5)
                    )),
                    &format_list(6),
                )?
                .write_with_format(3, 6, format!("{}월 이월금", month - 1), &format_list(3))?;
        }

        data_size.push(data_list.len() as u32);
        workbook.push_worksheet(worksheet);
    }

    // Save the file to disk.
    workbook.save("example/test.xlsx")?;

    run_shell_command()?;

    // send xlsx file to discord server
    send_discord_xlsx().await?;
    Ok(())
}
