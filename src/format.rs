use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder};

pub const BG_LAVENDER: Color = Color::RGB(0xCCC1DE); // lavender
pub const BG_GRAY: Color = Color::RGB(0xD8D8D8); // gray
pub const BATANGCHE: &str = "Batangche";
pub const NUM_FORMAT_STR: &str = "_-₩* #,##0_-;-₩* #,##0_-;_-₩* \" - \"_-;_-@";
pub const DATE_FORMAT_STR: &str = "mm\"월\" dd\"일\"";

pub fn format_list(idx: usize) -> Format {
    let FORMAT_LIST: [Format; 8] = [
        // [0] month
        Format::new()
            .set_align(FormatAlign::VerticalCenter)
            .set_background_color(BG_LAVENDER)
            .set_border(FormatBorder::Thin)
            .set_align(FormatAlign::Center)
            .set_font_name(BATANGCHE)
            .set_font_size(15)
            .set_bold(),
        // [1] schema
        Format::new()
            .set_align(FormatAlign::VerticalCenter)
            .set_align(FormatAlign::Center)
            .set_background_color(BG_LAVENDER)
            .set_border(FormatBorder::Thin)
            .set_font_name(BATANGCHE)
            .set_font_size(10)
            .set_bold(),
        // [2] text
        Format::new()
            .set_align(FormatAlign::VerticalCenter)
            .set_align(FormatAlign::Center)
            .set_border(FormatBorder::Thin)
            .set_font_name(BATANGCHE)
            .set_font_size(10),
        // [3] gray text
        Format::new()
            .set_align(FormatAlign::VerticalCenter)
            .set_align(FormatAlign::Center)
            .set_background_color(BG_GRAY)
            .set_border(FormatBorder::Thin)
            .set_font_name(BATANGCHE)
            .set_font_size(10),
        // [4] lavender text
        Format::new()
            .set_align(FormatAlign::VerticalCenter)
            .set_align(FormatAlign::Center)
            .set_background_color(BG_LAVENDER)
            .set_border(FormatBorder::Thin)
            .set_font_name(BATANGCHE)
            .set_font_size(10),
        // [5] num
        Format::new()
            .set_align(FormatAlign::VerticalCenter)
            .set_align(FormatAlign::Center)
            .set_border(FormatBorder::Thin)
            .set_font_name(BATANGCHE)
            .set_font_size(10)
            .set_num_format(NUM_FORMAT_STR),
        // [6] gray formula
        Format::new()
            .set_align(FormatAlign::VerticalCenter)
            .set_align(FormatAlign::Center)
            .set_background_color(BG_GRAY)
            .set_border(FormatBorder::Thin)
            .set_font_name(BATANGCHE)
            .set_font_size(10)
            .set_num_format(NUM_FORMAT_STR),
        // [7] lavender formula
        Format::new()
            .set_align(FormatAlign::VerticalCenter)
            .set_align(FormatAlign::Center)
            .set_background_color(BG_LAVENDER)
            .set_border(FormatBorder::Thin)
            .set_font_name(BATANGCHE)
            .set_font_size(10)
            .set_num_format(NUM_FORMAT_STR),
    ];

    FORMAT_LIST[idx].clone()
}
