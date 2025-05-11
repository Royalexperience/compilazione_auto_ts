use std::fs;

use internal_resources::{formats,utility_worksheet,calendario};
use rust_xlsxwriter::{Format, Workbook, Worksheet, XlsxBorder, XlsxError};
mod internal_resources;

fn main() -> Result<(), XlsxError> {

    //------Creazione worksheet------//
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    //------Utility per celle custom ------//
    utility_worksheet::custom_height_and_width_for_cells(worksheet);

    //------Stringhe Statiche inizio ------//
    utility_worksheet::build_static_strings_in_excel(worksheet);
    let calendario = calendario::Calendario::new();
    worksheet.write_string(1, 14, &format!("{} {}", calendario.mese, calendario.anno), &formats::times_new_roman_bold_underline())?;

    //-----Costruzione dei bordi-----//
    //worksheet.write_blank(10, 10, &formats::set_border_top())?;

    //-----Costruzione Dinamica del mese----//
    utility_worksheet::costruisci_gg_nome_gg(worksheet, &calendario);

    utility_worksheet::build_top_lines(worksheet);

    workbook.save("example1.xlsx")?;
    Ok(())
}