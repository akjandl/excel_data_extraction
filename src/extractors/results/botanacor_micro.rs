use calamine::{DataType, Range};

use crate::extractors::ExtractionManager;
use crate::excel_tools::{ColIndexer, PotentialSheet, Sheet, SheetExtractor, SheetSelector};

pub fn get_extractors() -> ExtractionManager {
    let mid_2020_validator = ("AgrBotMap", |ws: &Range<DataType>| {
        if let Some(dt) = ws.get_value((0, 3)) {
            match dt {
                DataType::String(s) => s == &"AgricorSampleName".to_string(),
                _ => false,
            }
        } else {
            false
        }
    });
    let early_2020_validator = ("Master List", |ws: &Range<DataType>| {
        if let Some(dt) = ws.get_value((0, 1)) {
            match dt {
                DataType::String(s) => s == &"Test Id".to_string(),
                _ => false,
            }
        } else {
            false
        }
    });
    let pre_test_id_validator = ("Master List", |ws: &Range<DataType>| {
        if let Some(dt) = ws.get_value((0, 1)) {
            match dt {
                DataType::String(s) => s.starts_with("Company Name"),
                _ => false,
            }
        } else {
            false
        }
    });

    let test_id_sheet = SheetExtractor::Multi(SheetSelector {
        col_names: vec!["Test Id"],
        potential_sheets: vec![
            PotentialSheet {
                sheet_name: "AgrBotMap",
                col_indexers: vec![ColIndexer::Index(3)],
                sheet_for_val: mid_2020_validator.0,
                validator: mid_2020_validator.1,
            },
            PotentialSheet {
                sheet_name: "Master List",
                col_indexers: vec![ColIndexer::Index(1)],
                sheet_for_val: early_2020_validator.0,
                validator: early_2020_validator.1,
            },
            PotentialSheet {
                sheet_name: "Master List",
                col_indexers: vec![ColIndexer::DefaultValue(DataType::String("NA".to_string()))],
                sheet_for_val: pre_test_id_validator.0,
                validator: pre_test_id_validator.1,
            },
        ],
    });

    let sample_info_sheet = SheetExtractor::Multi(SheetSelector {
        col_names: vec![
            "Company",
            "Sample Name",
            "Sample Type",
            "PBST wt (g)",
            "PBST vol (mL)",
            "Enrichment wt (g)",
            "Enrichment vol (mL)",
        ],
        potential_sheets: vec![
            PotentialSheet {
                sheet_name: "Master List",
                col_indexers: vec![
                    ColIndexer::Index(6),
                    ColIndexer::Index(8),
                    ColIndexer::Index(10),
                    ColIndexer::Index(11),
                    ColIndexer::Index(13),
                    ColIndexer::Index(12),
                    ColIndexer::Index(14),
                ],
                sheet_for_val: early_2020_validator.0,
                validator: early_2020_validator.1,
            },
            PotentialSheet {
                sheet_name: "Master List",
                col_indexers: vec![
                    ColIndexer::Index(2),
                    ColIndexer::Index(3),
                    ColIndexer::Index(5),
                    ColIndexer::Index(6),
                    ColIndexer::Index(8),
                    ColIndexer::Index(7),
                    ColIndexer::Index(9),
                ],
                sheet_for_val: pre_test_id_validator.0,
                validator: pre_test_id_validator.1,
            },
        ],
    });

    let tym_values = SheetExtractor::Single(Sheet {
        sheet_name: "TYM Values",
        col_names: vec!["TYM CFU Count", "TYM Dil. Plate", "TYM Reported CFU/g"],
        col_indexers: vec![
            ColIndexer::Index(4),
            ColIndexer::Index(5),
            ColIndexer::Index(6),
        ],
    });
    let tot_aerobic_values = SheetExtractor::Single(Sheet {
        sheet_name: "Total Aerobic",
        col_names: vec!["TA CFU Count", "TA Dil. Plate", "TA Reported CFU/g"],
        col_indexers: vec![
            ColIndexer::Index(4),
            ColIndexer::Index(5),
            ColIndexer::Index(6),
        ],
    });
    let tot_col_values = SheetExtractor::Single(Sheet {
        sheet_name: "Total Coliforms",
        col_names: vec![
            "Coliforms CFU Count",
            "Coliforms Dil. Plate",
            "Coliforms Reported CFU/g",
        ],
        col_indexers: vec![
            ColIndexer::Index(4),
            ColIndexer::Index(5),
            ColIndexer::Index(6),
        ],
    });

    ExtractionManager::RowGrain(vec![
        test_id_sheet,
        sample_info_sheet,
        tym_values,
        tot_aerobic_values,
        tot_col_values,
    ])
}
