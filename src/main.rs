use rust_xlsxwriter::*;
use std::error::Error;
use std::path::Path;
use transaction_manager::send_file::run_shell_command;
use transaction_manager::{cell_name, extract_tables_from_pdf};

fn main() -> Result<(), Box<dyn Error>> {
    let pdf_path = Path::new("example/account2.pdf");
    let mut table = extract_tables_from_pdf(pdf_path)?;

    // Create a new Excel file object.
    let mut workbook = Workbook::new();
    let format_date = Format::new().set_num_format("mm\"월\" dd\"일\"");
    let format_digit = Format::new().set_num_format("_-₩* #,##0_-;-₩* #,##0_-;_-₩* \" - \"_-;_-@");
    let mut row = 0;

    while let Some(data) = table.pop() {
        // Add a worksheet to the workbook.
        let sheet_name = data.date.month.to_string() + "월 정산서";
        let worksheet = match workbook.worksheet_from_name(&sheet_name) {
            Ok(sheet) => {
                row += 1;
                sheet
            }
            Err(_) => {
                row = 0;
                workbook
                    .add_worksheet()
                    .set_name(sheet_name)
                    .expect("Add new worksheet")
            }
        };

        let datetime = ExcelDateTime::from_ymd(data.date.year, data.date.month, data.date.day)?;

        // Write with format
        worksheet.write_with_format(6 + row, 0, datetime, &format_date)?;
        worksheet.set_column_width(0, 8.64)?;

        worksheet.write_with_format(6 + row, 3, data.cash_in, &format_digit)?;
        worksheet.set_column_width(3, 10.64)?;

        worksheet.write_with_format(6 + row, 4, data.cash_out, &format_digit)?;
        worksheet.set_column_width(4, 10.64)?;

        // worksheet.write_with_format(6 + row, 5, data.balance, &format_digit)?;

        worksheet.write_formula_with_format(
            6 + row,
            5,
            Formula::new(format!(
                "={} + {} - {}",
                cell_name(5 + row, 5),
                cell_name(6 + row, 3),
                cell_name(6 + row, 4)
            )),
            &format_digit,
        )?;
        worksheet.set_column_width(5, 11.64)?;
    }

    // Save the file to disk.
    workbook.save("example/test.xlsx")?;

    run_shell_command()?;

    Ok(())
}
