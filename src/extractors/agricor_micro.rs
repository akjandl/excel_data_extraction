use calamine::{DataType, Range};
use std::rc::Rc;

use crate::excel_tools::column_finders::header_match;
use crate::excel_tools::{ColIndexer, Sheet, SheetExtractor};

fn find_col(starts_with: &'static str) -> Rc<dyn Fn(&Range<DataType>, u32) -> Range<DataType>> {
    Rc::new(move |ws: &Range<DataType>, row_count: u32| {
        header_match(ws, starts_with, 0, row_count, None)
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
            ColIndexer::ColFindFunc(find_col("Test ID")),
            ColIndexer::ColFindFunc(find_col("LICENSE NAME")),
            ColIndexer::ColFindFunc(find_col("CUSTOMER LICENSE")),
            ColIndexer::ColFindFunc(find_col("SAMPLE NAME")),
            ColIndexer::ColFindFunc(find_col("SAMPLE TYPE")),
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
            ColIndexer::ColFindFunc(find_col("TYM Sample Weight")),
            ColIndexer::ColFindFunc(find_col("TYM Diluent Vol")),
            ColIndexer::ColFindFunc(find_col("Colony Count")),
            ColIndexer::ColFindFunc(find_col("Dilution Plate")),
            ColIndexer::ColFindFunc(find_col("Metrc Reported CFU")),
        ],
    });

    vec![master_list, tym_sheet]
}
