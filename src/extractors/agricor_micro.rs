use calamine::{DataType, Range};

use crate::excel_tools::column_finders::header_match;
use crate::excel_tools::{ColIndexer, PotentialSheet, Sheet, SheetExtractor, SheetSelector};

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

    vec![master_list]
}
