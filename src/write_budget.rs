use std::error::Error;

use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, Formula, Worksheet};

use crate::{
    cell_name,
    format::format_list,
    models::data::{BusinessType, VariantName},
};

pub fn write_business(
    worksheet: &mut Worksheet,
    business_type: BusinessType,
    row: u32,
    cnt: u32,
) -> Result<u32, Box<dyn Error>> {
    worksheet.merge_range(
        row,
        1,
        row + cnt - 1,
        1,
        business_type.variant_name(),
        &format_list(2)
            .set_font_size(12)
            .set_border(FormatBorder::Thin)
            .set_border_left(FormatBorder::Medium),
    )?;
    // // 개강파티
    // worksheet
    //     .write_row_with_format(
    //         row,
    //         2,
    //         ["개강파티", "", "", "", "", ""],
    //         &format_list(2)
    //             .set_font_size(12)
    //             .set_border(FormatBorder::Thin),
    //     )?
    //     .set_row_height(row, 27.8)?;
    // // 간식행사
    // worksheet
    //     .merge_range(
    //         row + 1,
    //         2,
    //         row + 2,
    //         2,
    //         "간식행사",
    //         &format_list(2)
    //             .set_font_size(12)
    //             .set_border(FormatBorder::Thin),
    //     )?
    //     .write_row_with_format(
    //         row + 1,
    //         3,
    //         ["중간고사", "", "", "", ""],
    //         &format_list(2)
    //             .set_font_size(12)
    //             .set_border(FormatBorder::Thin),
    //     )?
    //     .set_row_height(row + 1, 27.8)?
    //     .write_row_with_format(
    //         row + 2,
    //         3,
    //         ["기말고사", "", "", "", ""],
    //         &format_list(2)
    //             .set_font_size(12)
    //             .set_border(FormatBorder::Thin),
    //     )?
    //     .set_row_height(row + 2, 27.8)?;
    for r in row..row + cnt {
        worksheet
            .write_row_with_format(
                r,
                2,
                ["", "", "", "", "", ""],
                &format_list(2)
                    .set_font_size(12)
                    .set_border(FormatBorder::Thin),
            )?
            .set_row_height(r, 38)?;
        worksheet.write_with_format(
            r,
            8,
            "",
            &Format::new().set_border_left(FormatBorder::Medium),
        )?;
    }
    Ok(row + cnt)
}

pub fn budget(worksheet: &mut Worksheet, period: &String) -> Result<(), Box<dyn Error>> {
    let schema_format = Format::new()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_font_name("Batangche")
        .set_font_size(12)
        .set_background_color(Color::RGB(0xE5E0Ef));

    // set row height
    worksheet
        .set_row_height(0, 99)?
        .set_row_height(1, 7.5)?
        .set_row_height(2, 22.5)?
        .set_row_height(3, 7.5)?
        .set_row_height(4, 22.5)?
        .set_row_height(5, 22.5)?
        .set_row_height(6, 26.3)?
        .set_row_height(7, 15.8)?
        .set_row_height(8, 26.3)?
        .set_row_height(9, 27.8)?;

    // set column width
    worksheet
        .set_column_width(0, 1.64)?
        .set_column_width(1, 10.64)?
        .set_column_width(2, 14.36)?
        .set_column_width(3, 22.45)?
        .set_column_width(4, 15.64)?
        .set_column_width(5, 15.64)?
        .set_column_width(6, 35.91)?
        .set_column_width(7, 43.91)?;

    // Header
    worksheet
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
                    format!("{period} 예산안\n").as_str(),
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

    worksheet
    .merge_range(1, 1, 1, 7, "",&Format::new())?
    .merge_range(2, 1, 2, 7, "예산안 작성 전, 반드시 가이드라인 및 작성 예시를 참고해주세요. / 색칠된 칸은 입력하지 마세요. / 양식에 맞추어 작성해주시고, 예산안 원본도 첨부해주세요.", &format_list(2).set_font_size(12).set_bold().set_border(FormatBorder::Thin).set_border_color(Color::White))?;

    // 수입
    worksheet
        .merge_range(
            4,
            1,
            4,
            7,
            "수입",
            &schema_format.clone()
                .set_bold()
                .set_border(FormatBorder::Medium),
        )?
        .merge_range(
            5,
            1,
            5,
            3,
            "이월금",
            &schema_format
                .clone()
                .set_border(FormatBorder::Thin)
                .set_border_left(FormatBorder::Medium),
        )?
        .merge_range(
            5,
            4,
            5,
            6,
            "학생회비 납부",
            &schema_format
                .clone()
                .set_border(FormatBorder::Thin),
        )?
        .write_with_format(
            5,
            7,
            "계",
            &schema_format
                .clone()
                .set_font_size(12)
                .set_border(FormatBorder::Thin)
                .set_border_right(FormatBorder::Medium),
        )?
        .merge_range(
            6,
            1,
            6,
            3,
            "",
            &format_list(5)
                .set_font_size(12)
                .set_border(FormatBorder::Medium)
                .set_border_right(FormatBorder::Thin),
        )?
        .merge_range(
            6,
            4,
            6,
            6,
            "",
            &format_list(5)
                .set_font_size(12)
                .set_border(FormatBorder::Medium)
                .set_border_left(FormatBorder::Thin)
                .set_border_right(FormatBorder::Thin),
        )?
        .write_formula_with_format(
            6,
            4,
            Formula::new(
                "=SUMIF('1월 정산서'!C:C, \"학생회비 납부\", '1월 정산서'!D:D) + SUMIF('2월 정산서'!C:C, \"학생회비 납부\", '2월 정산서'!D:D) + SUMIF('3월 정산서'!C:C, \"학생회비 납부\", '3월 정산서'!D:D) + SUMIF('4월 정산서'!C:C, \"학생회비 납부\", '4월 정산서'!D:D) + SUMIF('5월 정산서'!C:C, \"학생회비 납부\", '5월 정산서'!D:D) + SUMIF('6월 정산서'!C:C, \"학생회비 납부\", '6월 정산서'!D:D)",
            ),&format_list(5)
            .set_font_size(12)
            .set_border(FormatBorder::Medium)
            .set_border_left(FormatBorder::Thin)
            .set_border_right(FormatBorder::Thin)
        )?
        .write_formula_with_format(
            6,
            7,
            Formula::new("=B7+E7"),
            &format_list(6)
            .set_font_size(12)
            .set_border(FormatBorder::Medium)
            .set_border_left(FormatBorder::Thin),
        )?;

    let row = write_business(worksheet, BusinessType::OngoingBusiness, 10, 2)?;
    let row = write_business(worksheet, BusinessType::GeneralBusiness, row, 4)?;
    let row = write_business(worksheet, BusinessType::PledgedBusiness, row, 2)?;

    // 지출
    worksheet
        .merge_range(
            8,
            1,
            8,
            7,
            "지출",
            &schema_format
                .clone()
                .set_bold()
                .set_border(FormatBorder::Medium),
        )?
        .write_with_format(
            9,
            1,
            "사업구분",
            &schema_format
                .clone()
                .set_border(FormatBorder::Thin)
                .set_border_left(FormatBorder::Medium)
                .set_border_bottom(FormatBorder::Medium),
        )?
        .write_with_format(
            9,
            2,
            "사업명",
            &schema_format
                .clone()
                .set_border(FormatBorder::Thin)
                .set_border_bottom(FormatBorder::Medium),
        )?
        .write_with_format(
            9,
            3,
            "지출 상세",
            &schema_format
                .clone()
                .set_border(FormatBorder::Thin)
                .set_border_bottom(FormatBorder::Medium),
        )?
        .write_with_format(
            9,
            4,
            format!("{}년도 예산", period[0..4].parse::<u32>()? - 1),
            &schema_format
                .clone()
                .set_border(FormatBorder::Thin)
                .set_border_bottom(FormatBorder::Medium),
        )?
        .write_with_format(
            9,
            5,
            format!("{}년도 예산", period[0..4].parse::<u32>()?),
            &schema_format
                .clone()
                .set_border(FormatBorder::Thin)
                .set_border_bottom(FormatBorder::Medium),
        )?
        .write_with_format(
            9,
            6,
            "산출 근거",
            &schema_format
                .clone()
                .set_border(FormatBorder::Thin)
                .set_border_bottom(FormatBorder::Medium),
        )?
        .write_with_format(
            9,
            7,
            "비고",
            &schema_format
                .clone()
                .set_border(FormatBorder::Thin)
                .set_border_right(FormatBorder::Medium)
                .set_border_bottom(FormatBorder::Medium),
        )?;

    worksheet
        .set_row_height(row, 38)?
        .merge_range(
            row,
            1,
            row,
            2,
            "예비비",
            &format_list(2)
                .set_font_size(12)
                .set_border(FormatBorder::Thin)
                .set_border_left(FormatBorder::Medium),
        )?
        .write_with_format(
            row,
            3,
            "잔액",
            &format_list(2)
                .set_font_size(12)
                .set_border(FormatBorder::Thin),
        )?
        .write_formula_with_format(
            row,
            4,
            Formula::new(format!(
                "={}-SUM({}:{})",
                cell_name(6, 6),
                cell_name(10, 4),
                cell_name(row - 1, 4)
            )),
            &format_list(5)
                .set_font_size(12)
                .set_border(FormatBorder::Thin),
        )?
        .write_formula_with_format(
            row,
            5,
            Formula::new(format!(
                "={}-SUM({}:{})",
                cell_name(6, 7),
                cell_name(10, 5),
                cell_name(row - 1, 5)
            )),
            &format_list(5)
                .set_font_size(12)
                .set_border(FormatBorder::Thin),
        )?
        .write_with_format(
            row,
            6,
            "",
            &format_list(2)
                .set_font_size(12)
                .set_border(FormatBorder::Thin),
        )?
        .write_with_format(
            row,
            7,
            "",
            &format_list(2)
                .set_font_size(12)
                .set_border(FormatBorder::Thin)
                .set_border_right(FormatBorder::Medium),
        )?
        .set_row_height(row + 1, 35)?
        .merge_range(
            row + 1,
            1,
            row + 1,
            4,
            "계",
            &format_list(3)
                .set_font_size(12)
                .set_border(FormatBorder::Medium)
                .set_border_right(FormatBorder::Thin),
        )?
        .write_formula_with_format(
            row + 1,
            5,
            Formula::new(format!("=SUM({}:{})", cell_name(10, 5), cell_name(row, 5))),
            &format_list(6)
                .set_font_size(12)
                .set_border(FormatBorder::Medium)
                .set_border_left(FormatBorder::Thin)
                .set_border_right(FormatBorder::Thin),
        )?
        .write_with_format(
            row + 1,
            6,
            "",
            &format_list(3)
                .set_font_size(12)
                .set_border(FormatBorder::Medium)
                .set_border_left(FormatBorder::Thin)
                .set_border_right(FormatBorder::Thin),
        )?
        .write_with_format(
            row + 1,
            7,
            "",
            &format_list(3)
                .set_font_size(12)
                .set_border(FormatBorder::Medium)
                .set_border_left(FormatBorder::Thin),
        )?;

    Ok(())
}
