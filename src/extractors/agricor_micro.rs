use calamine::{DataType, Range};

use crate::excel_tools::column_finders::header_match;
use crate::excel_tools::{ColIndexer, Sheet, SheetExtractor};

// fn get_ml_header(
//     starts_with: &'static str,
// ) -> fn(&Range<DataType>, u32) -> Range<DataType> {
//     |ws: &Range<DataType>, cols: u32|  header_match(ws, &starts_with, 0, cols, None)
// }

fn find_test_id(ws: &Range<DataType>, row_count: u32) -> Range<DataType> {
    header_match(ws, "Test ID", 0, row_count, None) 
}

fn find_customer(ws: &Range<DataType>, row_count: u32) -> Range<DataType> {
    header_match(ws, "LICENSE NAME", 0, row_count, None) 
}

fn find_license(ws: &Range<DataType>, row_count: u32) -> Range<DataType> {
    header_match(ws, "CUSTOMER LICENSE", 0, row_count, None) 
}

fn find_sample_name(ws: &Range<DataType>, row_count: u32) -> Range<DataType> {
    header_match(ws, "SAMPLE NAME", 0, row_count, None) 
}

fn find_sample_type(ws: &Range<DataType>, row_count: u32) -> Range<DataType> {
    header_match(ws, "SAMPLE TYPE", 0, row_count, None) 
}

fn find_tym_col_count(ws: &Range<DataType>, row_count: u32) -> Range<DataType> {
    header_match(ws, "Colony Count", 0, row_count, None) 
}

fn find_tym_dilution(ws: &Range<DataType>, row_count: u32) -> Range<DataType> {
    header_match(ws, "Dilution Plate", 0, row_count, None) 
}

fn find_tym_cfu(ws: &Range<DataType>, row_count: u32) -> Range<DataType> {
    header_match(ws, "Metrc Reported CFU", 0, row_count, None) 
}

fn find_spl_wt(ws: &Range<DataType>, row_count: u32) -> Range<DataType> {
    header_match(ws, "TYM Sample Weight", 0, row_count, None) 
}

fn find_dil_vol(ws: &Range<DataType>, row_count: u32) -> Range<DataType> {
    header_match(ws, "TYM Diluent Vol", 0, row_count, None) 
}

pub fn get_extractors() -> Vec<SheetExtractor> {
    let master_list = SheetExtractor::Single(Sheet {
        sheet_name: "Master List",
        col_names: vec![
            "Test ID",
            "Company",
            "Company License",
            "Sample Name",
            "Sample Type",
        ],
        col_indexers: vec![
            ColIndexer::ColFindFunc(find_test_id),
            ColIndexer::ColFindFunc(find_customer),
            ColIndexer::ColFindFunc(find_license),
            ColIndexer::ColFindFunc(find_sample_name),
            ColIndexer::ColFindFunc(find_sample_type),
        ],
    });

    let tym_sheet = SheetExtractor::Single(Sheet {
        sheet_name: "TYM Values",
        col_names: vec![
            "Sample Weight (g)",
            "Diluent Vol (mL)", 
            "Colony Count",
            "Dilution Plate",
            "Reported CFU/g"
        ],
        col_indexers: vec![
            ColIndexer::ColFindFunc(find_spl_wt),
            ColIndexer::ColFindFunc(find_dil_vol),
            ColIndexer::ColFindFunc(find_tym_col_count),
            ColIndexer::ColFindFunc(find_tym_dilution),
            ColIndexer::ColFindFunc(find_tym_cfu),
        ]
    });

    vec![master_list, tym_sheet]
}
