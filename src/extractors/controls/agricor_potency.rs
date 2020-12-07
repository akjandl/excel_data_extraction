use calamine::{DataType, Range};

use crate::excel_tools::{ColIndexer, PotentialSheet, Sheet, SheetExtractor, SheetSelector};
use crate::extractors::ExtractionManager;

pub fn get_extractors() -> ExtractionManager {
    let late_2020_validator = ("System Suitability", |ws: &Range<DataType>| {
        if let Some(dt) = ws.get_value((29, 0)) {
            match dt {
                DataType::String(s) => s == &"Bracketing Standard 1".to_string(),
                _ => false,
            }
        } else {
            false
        }
    });

    let early_2020_validator = ("System Suitability", |ws: &Range<DataType>| {
        if let Some(dt) = ws.get_value((34, 0)) {
            match dt {
                DataType::String(s) => s == &"Bracketing Standard 1".to_string(),
                _ => false,
            }
        } else {
            false
        }
    });

    let sys_suit_common = SheetExtractor::Single(Sheet {
        sheet_name: "System Suitability",
        col_names: vec![
            "Date",
            "IS Response",
            "Check Std. Low - CBD",
            "Check Std. Low - d9THC",
            "Check Std. High - CBD",
            "Check Std. High - d9THC",
        ],
        col_indexers: vec![
            ColIndexer::CellValue { row: 1, col: 3 },
            ColIndexer::CellValue { row: 6, col: 1 },
            ColIndexer::CellValue { row: 19, col: 1 },
            ColIndexer::CellValue { row: 20, col: 1 },
            ColIndexer::CellValue { row: 24, col: 1 },
            ColIndexer::CellValue { row: 25, col: 1 },
        ],
    });

    let sys_suit_variable = SheetExtractor::Multi(SheetSelector {
        col_names: vec![
            "Placebo - CBD",
            "Placebo - d9THC",
            "Bracketing Std. 1 - CBD",
            "Bracketing Std. 2 - CBD",
            "Bracketing Std. 3 - CBD",
            "Bracketing Std. 4 - CBD",
            "Bracketing Std. 5 - CBD",
            "Bracketing Std. 6 - CBD",
            "Bracketing Std. 7 - CBD",
            "Bracketing Std. 8 - CBD",
            "Bracketing Std. 9 - CBD",
            "Bracketing Std. 10 - CBD",
            "Bracketing Std. 1 - d9THC",
            "Bracketing Std. 2 - d9THC",
            "Bracketing Std. 3 - d9THC",
            "Bracketing Std. 4 - d9THC",
            "Bracketing Std. 5 - d9THC",
            "Bracketing Std. 6 - d9THC",
            "Bracketing Std. 7 - d9THC",
            "Bracketing Std. 8 - d9THC",
            "Bracketing Std. 9 - d9THC",
            "Bracketing Std. 10 - d9THC",
            "Spiked Surrogate 1 - CBD",
            "Spiked Surrogate 2 - CBD",
            "Spiked Surrogate 3 - CBD",
            "Spiked Surrogate 4 - CBD",
            "Spiked Surrogate 5 - CBD",
            "Spiked Surrogate 6 - CBD",
            "Spiked Surrogate 7 - CBD",
            "Spiked Surrogate 8 - CBD",
            "Spiked Surrogate 9 - CBD",
            "Spiked Surrogate 10 - CBD",
            "Spiked Surrogate 11 - CBD",
            "Spiked Surrogate 12 - CBD",
            "Spiked Surrogate 1 - d9THC",
            "Spiked Surrogate 2 - d9THC",
            "Spiked Surrogate 3 - d9THC",
            "Spiked Surrogate 4 - d9THC",
            "Spiked Surrogate 5 - d9THC",
            "Spiked Surrogate 6 - d9THC",
            "Spiked Surrogate 7 - d9THC",
            "Spiked Surrogate 8 - d9THC",
            "Spiked Surrogate 9 - d9THC",
            "Spiked Surrogate 10 - d9THC",
            "Spiked Surrogate 11 - d9THC",
            "Spiked Surrogate 12 - d9THC",
            "Primary Standard 1 - CBD",
            "Primary Standard 2 - CBD",
            "Primary Standard 3 - CBD",
            "Primary Standard 4 - CBD",
            "Primary Standard 5 - CBD",
            "Primary Standard 1 - d9THC",
            "Primary Standard 2 - d9THC",
            "Primary Standard 3 - d9THC",
            "Primary Standard 4 - d9THC",
            "Primary Standard 5 - d9THC",
        ],
        potential_sheets: vec![
            PotentialSheet {
                sheet_for_val: late_2020_validator.0,
                validator: late_2020_validator.1,
                sheet_name: "System Suitability",
                col_indexers: vec![
                    // placebos
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    // CBD bracketing standards
                    ColIndexer::CellValue{row: 29, col: 1},
                    ColIndexer::CellValue{row: 30, col: 1},
                    ColIndexer::CellValue{row: 31, col: 1},
                    ColIndexer::CellValue{row: 32, col: 1},
                    ColIndexer::CellValue{row: 33, col: 1},
                    ColIndexer::CellValue{row: 34, col: 1},
                    ColIndexer::CellValue{row: 35, col: 1},
                    ColIndexer::CellValue{row: 36, col: 1},
                    ColIndexer::CellValue{row: 37, col: 1},
                    ColIndexer::CellValue{row: 38, col: 1},
                    // d9THC bracketing standards
                    ColIndexer::CellValue{row: 29, col: 2},
                    ColIndexer::CellValue{row: 30, col: 2},
                    ColIndexer::CellValue{row: 31, col: 2},
                    ColIndexer::CellValue{row: 32, col: 2},
                    ColIndexer::CellValue{row: 33, col: 2},
                    ColIndexer::CellValue{row: 34, col: 2},
                    ColIndexer::CellValue{row: 35, col: 2},
                    ColIndexer::CellValue{row: 36, col: 2},
                    ColIndexer::CellValue{row: 37, col: 2},
                    ColIndexer::CellValue{row: 38, col: 2},
                    // CBD spiked surrogates
                    ColIndexer::CellValue{row: 42, col: 1},
                    ColIndexer::CellValue{row: 43, col: 1},
                    ColIndexer::CellValue{row: 44, col: 1},
                    ColIndexer::CellValue{row: 45, col: 1},
                    ColIndexer::CellValue{row: 46, col: 1},
                    ColIndexer::CellValue{row: 47, col: 1},
                    ColIndexer::CellValue{row: 48, col: 1},
                    ColIndexer::CellValue{row: 49, col: 1},
                    ColIndexer::CellValue{row: 50, col: 1},
                    ColIndexer::CellValue{row: 51, col: 1},
                    ColIndexer::CellValue{row: 52, col: 1},
                    ColIndexer::CellValue{row: 53, col: 1},
                    // d9THC spiked surrogates
                    ColIndexer::CellValue{row: 42, col: 2},
                    ColIndexer::CellValue{row: 43, col: 2},
                    ColIndexer::CellValue{row: 44, col: 2},
                    ColIndexer::CellValue{row: 45, col: 2},
                    ColIndexer::CellValue{row: 46, col: 2},
                    ColIndexer::CellValue{row: 47, col: 2},
                    ColIndexer::CellValue{row: 48, col: 2},
                    ColIndexer::CellValue{row: 49, col: 2},
                    ColIndexer::CellValue{row: 50, col: 2},
                    ColIndexer::CellValue{row: 51, col: 2},
                    ColIndexer::CellValue{row: 52, col: 2},
                    ColIndexer::CellValue{row: 53, col: 2},
                    // CBD primary standards
                    ColIndexer::CellValue{row: 75, col: 1},
                    ColIndexer::CellValue{row: 76, col: 1},
                    ColIndexer::CellValue{row: 77, col: 1},
                    ColIndexer::CellValue{row: 78, col: 1},
                    ColIndexer::CellValue{row: 79, col: 1},
                    // d9THC primary standards
                    ColIndexer::CellValue{row: 75, col: 2},
                    ColIndexer::CellValue{row: 76, col: 2},
                    ColIndexer::CellValue{row: 77, col: 2},
                    ColIndexer::CellValue{row: 78, col: 2},
                    ColIndexer::CellValue{row: 79, col: 2},
                ],
            },
            PotentialSheet {
                sheet_for_val: early_2020_validator.0,
                validator: early_2020_validator.1,
                sheet_name: "System Suitability",
                col_indexers: vec![
                    // placebos
                    ColIndexer::CellValue{row: 29, col: 1},
                    ColIndexer::CellValue{row: 30, col: 1},
                    // CBD bracketing standards
                    ColIndexer::CellValue{row: 34, col: 1},
                    ColIndexer::CellValue{row: 35, col: 1},
                    ColIndexer::CellValue{row: 36, col: 1},
                    ColIndexer::CellValue{row: 37, col: 1},
                    ColIndexer::CellValue{row: 38, col: 1},
                    ColIndexer::CellValue{row: 39, col: 1},
                    ColIndexer::CellValue{row: 40, col: 1},
                    ColIndexer::CellValue{row: 41, col: 1},
                    ColIndexer::CellValue{row: 42, col: 1},
                    ColIndexer::CellValue{row: 43, col: 1},
                    // d9THC bracketing standards
                    ColIndexer::CellValue{row: 34, col: 2},
                    ColIndexer::CellValue{row: 35, col: 2},
                    ColIndexer::CellValue{row: 36, col: 2},
                    ColIndexer::CellValue{row: 37, col: 2},
                    ColIndexer::CellValue{row: 38, col: 2},
                    ColIndexer::CellValue{row: 39, col: 2},
                    ColIndexer::CellValue{row: 40, col: 2},
                    ColIndexer::CellValue{row: 41, col: 2},
                    ColIndexer::CellValue{row: 42, col: 2},
                    ColIndexer::CellValue{row: 43, col: 2},
                    // CBD spiked surrogates
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    // d9THC spiked surrogates
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    // CBD primary standards
                    ColIndexer::CellValue{row: 63, col: 1},
                    ColIndexer::CellValue{row: 64, col: 1},
                    ColIndexer::CellValue{row: 65, col: 1},
                    ColIndexer::CellValue{row: 66, col: 1},
                    ColIndexer::CellValue{row: 67, col: 1},
                    // d9THC primary standards
                    ColIndexer::CellValue{row: 63, col: 2},
                    ColIndexer::CellValue{row: 64, col: 2},
                    ColIndexer::CellValue{row: 65, col: 2},
                    ColIndexer::CellValue{row: 66, col: 2},
                    ColIndexer::CellValue{row: 67, col: 2},
                ],
            }
        ],
    });

    ExtractionManager::FileGrain(vec![sys_suit_common, sys_suit_variable])
}
