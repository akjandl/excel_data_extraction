use calamine::{DataType, Range};
use std::rc::Rc;

use crate::excel_tools::column_finders::{header_match, MatchMethod};
use crate::excel_tools::{ColIndexer, Sheet, SheetExtractor};

fn header_starts_with(
    starts_with: &'static str,
) -> Rc<dyn Fn(&Range<DataType>, u32) -> Range<DataType>> {
    Rc::new(move |ws: &Range<DataType>, row_count: u32| {
        header_match(ws, MatchMethod::StartsWith(starts_with), 0, row_count, None)
    })
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
            ColIndexer::ColFindFunc(header_starts_with("Test ID")),
            ColIndexer::ColFindFunc(header_starts_with("LICENSE NAME")),
            ColIndexer::ColFindFunc(header_starts_with("CUSTOMER LICENSE")),
            ColIndexer::ColFindFunc(header_starts_with("SAMPLE NAME")),
            ColIndexer::ColFindFunc(header_starts_with("SAMPLE TYPE")),
        ],
    });

    let tym_sheet = SheetExtractor::Single(Sheet {
        sheet_name: "TYM Values",
        col_names: vec![
            "Sample Weight (g)",
            "Diluent Vol (mL)",
            "Colony Count",
            "Dilution Plate",
            "Reported CFU/g",
        ],
        col_indexers: vec![
            ColIndexer::ColFindFunc(header_starts_with("TYM Sample Weight")),
            ColIndexer::ColFindFunc(header_starts_with("TYM Diluent Vol")),
            ColIndexer::ColFindFunc(header_starts_with("Colony Count")),
            ColIndexer::ColFindFunc(header_starts_with("Dilution Plate")),
            ColIndexer::ColFindFunc(header_starts_with("Metrc Reported CFU")),
        ],
    });

    vec![master_list, tym_sheet]
}
