use rust_xlsxwriter::*;
use std::error::Error;
use std::path::Path;
use transaction_manager::discord_message::send_discord_xlsx;
use transaction_manager::format::format_list;
use transaction_manager::write_account::account;
use transaction_manager::write_budget::budget;
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

    let mut worksheets = Vec::with_capacity(7);
    let month_data_list = separate_data(table)?;
    let mut data_size = Vec::with_capacity(month_data_list.len()); // 작월 이월금을 가져오기 위해 사용

    //
    let period = format!(
        "{}년도 제{}회기",
        month_data_list[0].1[0].date.year,
        match month_data_list[0].0 {
            // 이 부분은 개선의 필요가 있을 듯
            6 => 2,
            1 => 1,
            _ => 0,
        }
    );

    for (month, data_list) in month_data_list.iter() {
        let sheet_name = month.to_string() + "월 정산서";
        let mut worksheet = Worksheet::new();
        sheet_template(&mut worksheet, sheet_name.as_str())?;

        // write data
        write_data_in_sheet(&mut worksheet, data_list)?;

        // write schema formula
        let len = data_list.len() as u32;
        worksheet
            // 수입
            .write_formula_with_format(
                1,
                3,
                Formula::new(format!("={}", cell_name(7 + len, 3))),
                &format_list(6),
            )?
            // 지출
            .write_formula_with_format(
                2,
                3,
                Formula::new(format!("={}", cell_name(7 + len, 4))),
                &format_list(6),
            )?;
        // 이월금
        if worksheets.is_empty() {
            worksheet
                .write_formula_with_format(
                    3,
                    3,
                    Formula::new(format!("='{} 예산안'!B7", period)),
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
                        cell_name(7 + data_size.last().unwrap(), 5)
                    )),
                    &format_list(6),
                )?
                .write_with_format(3, 6, format!("{}월 이월금", month - 1), &format_list(3))?;
        }

        data_size.push(data_list.len() as u32);
        worksheets.push(worksheet);
    }

    // {}년도 제{}회기 예산안
    let worksheet1 = workbook
        .add_worksheet()
        .set_name(format!("{} 예산안", period))?;

    // budget
    budget(worksheet1, &period)?;

    // {}년도 제{}회기 정산서
    let worksheet2 = workbook
        .add_worksheet()
        .set_name(format!("{} 정산서", period))?;

    // account
    account(worksheet2, &period)?;

    for worksheet in worksheets.into_iter() {
        workbook.push_worksheet(worksheet);
    }

    // Save the file to disk.
    workbook.save("example/test.xlsx")?;

    // run_shell_command()?;

    // send xlsx file to discord server
    send_discord_xlsx().await?;
    Ok(())
}
