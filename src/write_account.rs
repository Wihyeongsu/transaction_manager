use std::error::Error;

use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, Formula, Worksheet};

use crate::{cell_name, format::format_list};

pub fn account(worksheet: &mut Worksheet, period: &String) -> Result<(), Box<dyn Error>> {
    let schema_format = Format::new()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_font_name("Batangche")
        .set_font_size(12)
        .set_background_color(Color::RGB(0xE5E0Ef));

    // set column width
    worksheet
        .set_column_width(0, 1.64)?
        .set_column_width(1, 10.64)?
        .set_column_width(2, 14.36)?
        .set_column_width(3, 14.36)?
        .set_column_width(4, 15.64)?
        .set_column_width(5, 15.64)?
        .set_column_width(6, 46.45)?
        .set_column_width(7, 43.91)?;

    // Header
    worksheet
        .set_row_height(0, 99)?
        .merge_range(
            0,
            1,
            0,
            7,
            "",
            &Format::new()
                .set_align(FormatAlign::Center)
                .set_align(FormatAlign::VerticalCenter)
                .set_font_name("Arial")
                .set_background_color(Color::RGB(0xFCD5B6))
                .set_border(FormatBorder::Medium),
        )?
        .write_rich_string_with_format(
            0,
            1,
            &[
                (
                    &Format::new()
                        .set_font_size(30)
                        .set_font_name("새굴림")
                        .set_font_name("Arial")
                        .set_bold(),
                    format!("{period} 재정감사 정산서\n").as_str(),
                ),
                (
                    &Format::new()
                        .set_font_size(16)
                        .set_font_name("새굴림")
                        .set_font_name("Arial")
                        .set_bold(),
                    "(OOOO대학 OOOO학과 제OO대 OOOO학생회)\n",
                ),
                (
                    &Format::new()
                        .set_font_size(16)
                        .set_font_name("새굴림")
                        .set_font_name("Arial")
                        .set_bold(),
                    "(출범일 - yyyy.mm.dd~yyyy.mm.dd)",
                ),
            ],
            &Format::new()
                .set_align(FormatAlign::Center)
                .set_align(FormatAlign::VerticalCenter)
                .set_background_color(Color::RGB(0xFCD5B6))
                .set_border(FormatBorder::Medium),
        )?;

    worksheet.set_row_height(1, 22.5)?.merge_range(
        1,
        1,
        1,
        7,
        "주황색 칸의 학생회명 및 월별 영수증 번호만 적어주세요.",
        &format_list(2)
            .set_font_size(12)
            .set_bold()
            .set_border(FormatBorder::Medium),
    )?;

    worksheet
        .set_row_height(2, 26.3)?
        .write_row_with_format(
            2,
            1,
            [
                "월",
                "수입",
                "지출",
                "이월금",
                "총잔액",
                "영수증 번호",
                "비고",
            ],
            &schema_format
                .clone()
                .set_border(FormatBorder::Thin)
                .set_border_top(FormatBorder::Medium)
                .set_border_bottom(FormatBorder::Medium),
        )?
        .write_with_format(
            2,
            0,
            "",
            &Format::new().set_border_right(FormatBorder::Medium),
        )?
        .write_with_format(
            2,
            8,
            "",
            &Format::new().set_border_left(FormatBorder::Medium),
        )?;

    worksheet
        .set_row_height(3, 27.8)?
        .write_with_format(
            3,
            1,
            "",
            &format_list(3)
                .set_font_size(12)
                .set_border(FormatBorder::Thin)
                .set_border_left(FormatBorder::Medium)
                .set_border_top(FormatBorder::Medium),
        )?
        .write_with_format(
            3,
            2,
            "",
            &format_list(3)
                .set_font_size(12)
                .set_border(FormatBorder::Thin)
                .set_border_top(FormatBorder::Medium)
                .set_border_right(FormatBorder::None),
        )?
        .write_with_format(
            3,
            3,
            "",
            &format_list(3)
                .set_font_size(12)
                .set_border(FormatBorder::Thin)
                .set_border_left(FormatBorder::None)
                .set_border_top(FormatBorder::Medium),
        )?
        .write_formula_with_format(
            3,
            4,
            Formula::new(format!("='{} 예산안'!B7", period)),
            &format_list(6)
                .set_font_size(12)
                .set_border(FormatBorder::Thin)
                .set_border_top(FormatBorder::Medium),
        )?
        .write_formula_with_format(
            3,
            5,
            Formula::new(format!("=E4")),
            &format_list(6)
                .set_font_size(12)
                .set_border(FormatBorder::Thin)
                .set_border_top(FormatBorder::Medium),
        )?
        .write_with_format(
            3,
            6,
            "",
            &format_list(2)
                .set_font_size(12)
                .set_border(FormatBorder::Thin)
                .set_border_top(FormatBorder::Medium),
        )?
        .write_with_format(
            3,
            7,
            "",
            &format_list(2)
                .set_font_size(12)
                .set_border(FormatBorder::Thin)
                .set_border_top(FormatBorder::Medium)
                .set_border_right(FormatBorder::Medium),
        )?;
    let mut row = 4;

    let (s, e) = match period.chars().nth(8).unwrap().to_string().parse::<u32>()? {
        1 => (1, 6),
        2 => (6, 12),
        _ => (0, 0),
    };

    for month in s..=e {
        worksheet
            .set_row_height(row, 27.8)?
            .write_with_format(
                row,
                1,
                format!("{}월", month),
                &format_list(3)
                    .set_font_size(12)
                    .set_border(FormatBorder::Thin)
                    .set_border_left(FormatBorder::Medium),
            )?
            .write_row_with_format(
                row,
                2,
                [
                    Formula::new(format!("='{}월 정산서'!D2", month)),
                    Formula::new(format!("='{}월 정산서'!D3", month)),
                    Formula::new(format!("={}", cell_name(row - 1, 5))),
                    Formula::new(format!(
                        "=SUM({}+{}-{})",
                        cell_name(row, 4),
                        cell_name(row, 2),
                        cell_name(row, 3)
                    )),
                ],
                &format_list(6)
                    .set_font_size(12)
                    .set_border(FormatBorder::Thin),
            )?
            .write_row_with_format(
                row,
                6,
                ["", ""],
                &format_list(2)
                    .set_font_size(12)
                    .set_border(FormatBorder::Thin),
            )?
            .write_with_format(
                row,
                8,
                "",
                &Format::new().set_border_left(FormatBorder::Medium),
            )?;

        row += 1;
    }

    // 계
    worksheet
        .set_row_height(row, 27)?
        .write_with_format(
            row,
            1,
            "계",
            &format_list(3)
                .set_font_size(12)
                .set_border(FormatBorder::Medium)
                .set_border_right(FormatBorder::Thin),
        )?
        .write_row_with_format(
            row,
            2,
            [
                Formula::new(format!(
                    "=SUM({}:{})",
                    cell_name(3, 2),
                    cell_name(row - 1, 2)
                )),
                Formula::new(format!(
                    "=SUM({}:{})",
                    cell_name(3, 3),
                    cell_name(row - 1, 3)
                )),
            ],
            &format_list(6)
                .set_font_size(12)
                .set_border(FormatBorder::Medium)
                .set_border_left(FormatBorder::Thin)
                .set_border_right(FormatBorder::Thin),
        )?
        .write_with_format(
            row,
            4,
            "-",
            &format_list(3)
                .set_font_size(12)
                .set_border(FormatBorder::Medium)
                .set_border_left(FormatBorder::Thin)
                .set_border_right(FormatBorder::Thin),
        )?
        .write_formula_with_format(
            row,
            5,
            Formula::new(format!("={}", cell_name(row - 1, 5))),
            &&format_list(6)
                .set_font_size(12)
                .set_border(FormatBorder::Medium)
                .set_border_left(FormatBorder::Thin)
                .set_border_right(FormatBorder::Thin),
        )?
        .write_row_with_format(
            row,
            6,
            ["", ""],
            &format_list(2)
                .set_font_size(12)
                .set_border(FormatBorder::Medium)
                .set_border_left(FormatBorder::Thin)
                .set_border_right(FormatBorder::Thin),
        )?
        .write_with_format(
            row,
            8,
            "",
            &Format::new().set_border_left(FormatBorder::Medium),
        )?;

    Ok(())
}
