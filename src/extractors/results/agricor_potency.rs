use calamine::{DataType, Range};
use std::rc::Rc;

use crate::extractors::ExtractionManager;
use crate::excel_tools::column_finders::{header_match, MatchMethod};
use crate::excel_tools::{ColIndexer, Sheet, SheetExtractor};

fn find_col(starts_with: &'static str) -> Rc<dyn Fn(&Range<DataType>, u32) -> Range<DataType>> {
    Rc::new(move |ws: &Range<DataType>, row_count: u32| {
        header_match(ws, MatchMethod::StartsWith(starts_with), 0, row_count, None)
    })
}

fn header_exact(match_str: &'static str) -> Rc<dyn Fn(&Range<DataType>, u32) -> Range<DataType>> {
    Rc::new(move |ws: &Range<DataType>, row_count: u32| {
        header_match(ws, MatchMethod::Exact(match_str), 0, row_count, None)
    })
}

pub fn get_extractors() -> ExtractionManager {
    let master_list = SheetExtractor::Single(Sheet {
        sheet_name: "Master List",
        col_names: vec![
            "Test ID",
            "Company",
            "Company License",
            "Manifest",
            "Sample Name",
            "Sample Type",
        ],
        col_indexers: vec![
            ColIndexer::ColFindFunc(find_col("Test ID")),
            ColIndexer::ColFindFunc(find_col("Customer License Name")),
            ColIndexer::ColFindFunc(find_col("Customer License Number")),
            ColIndexer::ColFindFunc(find_col("Manifest")),
            ColIndexer::ColFindFunc(find_col("Sample Name")),
            ColIndexer::ColFindFunc(find_col("Type")),
        ],
    });
    // TODO: add the other analytes here someday
    let results = SheetExtractor::Single(Sheet {
        sheet_name: "Sample Data",
        col_names: vec!["CBDa", "CBDVa", "CBDV"],
        col_indexers: vec![
            ColIndexer::ColFindFunc(header_exact("CBDa")),
            ColIndexer::ColFindFunc(header_exact("CBDVa")),
            ColIndexer::ColFindFunc(header_exact("CBDV")),
        ],
    });
    let limit_of_quants = SheetExtractor::Single(Sheet {
        sheet_name: "LOQ Summary",
        col_names: vec!["CBDa LLOQ", "CBDVa LLOQ", "CBDV LLOQ"],
        col_indexers: vec![
            ColIndexer::ColFindFunc(header_exact("CBDa")),
            ColIndexer::ColFindFunc(header_exact("CBDVa")),
            ColIndexer::ColFindFunc(header_exact("CBDV")),
        ],
    });
    let sample_prep = SheetExtractor::Single(Sheet {
        sheet_name: "Sample Prep Form",
        col_names: vec!["HPLC", "Start Date"],
        col_indexers: vec![ColIndexer::CellValue{row: 7, col: 1}, ColIndexer::CellValue{row: 0, col: 4}],
    });

    ExtractionManager::RowGrain(vec![master_list, results, limit_of_quants, sample_prep])
}
