use calamine::{DataType, Range, Reader, Xlsx};
use std::io::{Read, Seek};

pub fn find_active_rows(ws: &Range<DataType>, search_col: u32, last_row: Option<u32>) -> Vec<u32> {
    let bottom_row = match last_row {
        Some(num) => num,
        None => 200,
    };
    let mut active_rows: Vec<u32> = vec![];
    for row_num in 1..=bottom_row {
        let val = ws.get_value((row_num, search_col));
        match val {
            Some(dt) => match dt {
                DataType::Error(_e) => continue,
                DataType::Empty => continue,
                DataType::String(s) => {
                    if s == "" {
                        continue;
                    } else {
                        active_rows.push(row_num)
                    }
                }
                _ => active_rows.push(row_num),
            },
            None => continue,
        }
    }

    active_rows
}

pub fn extract_column(
    range: &Range<DataType>,
    col_indexer: &ColIndexer,
    max_row: u32,
) -> Range<DataType> {
    match col_indexer {
        ColIndexer::Index(i) => range.range((1, *i), (max_row, *i)),
        ColIndexer::DefaultValue(dt) => {
            let mut new_range = Range::new((0, 0), (max_row, 0));
            (0..max_row).for_each(|i| new_range.set_value((i, 0), dt.clone()));
            new_range
        }
    }
}

pub fn extract_sheet_columns(
    range: &Range<DataType>,
    sheet: &Sheet,
    max_row: u32,
) -> Vec<Range<DataType>> {
    let mut cols = vec![];
    for indexer in sheet.col_indexers.iter() {
        cols.push(extract_column(range, indexer, max_row))
    }
    cols
}

pub fn rows_from_cols(cols: Vec<Range<DataType>>, active_rows: Vec<u32>) -> Vec<Vec<DataType>> {
    let mut rows_data = vec![];
    for row in active_rows.iter() {
        let mut cur_row = vec![];
        for col in cols.iter() {
            cur_row.push(
                // Get range values by relative position in column.
                // Subtract 1 from the row since the `active_rows` should always
                // be offset by 1.
                col.get(((*row - 1) as usize, 0 as usize)).unwrap().clone(),
                // col.get((*row as usize, col.start().unwrap().1 as usize))
            );
        }
        rows_data.push(cur_row);
    }
    rows_data
}

pub fn make_header(sheets: &[SheetExtractor]) -> Vec<&'static str> {
    sheets
        .iter()
        .flat_map(|s| match s {
            SheetExtractor::Multi(s) => s.col_names.clone(),
            SheetExtractor::Single(s) => s.col_names.clone(),
        })
        .collect()
}

#[derive(Clone)]
pub enum ColIndexer {
    Index(u32),
    DefaultValue(DataType),
}

#[derive(Clone)]
pub struct Sheet {
    pub sheet_name: &'static str,
    pub col_names: Vec<&'static str>,
    pub col_indexers: Vec<ColIndexer>,
}

pub enum SheetExtractor {
    Single(Sheet),
    Multi(SheetSelector),
}

struct PotentialSheet {
    sheet_name: &'static str,
    col_indexers: Vec<ColIndexer>,
    sheet_for_val: &'static str,
    validator: fn(ws: &Range<DataType>) -> bool,
}

pub struct SheetSelector {
    col_names: Vec<&'static str>,
    potential_sheets: Vec<PotentialSheet>,
}

fn sheet_from_selector<RS: Read + Seek>(
    wb: &mut Xlsx<RS>,
    selector: &SheetSelector,
) -> Option<Sheet> {
    for p_sheet in selector.potential_sheets.iter() {
        if let Some(Ok(ws)) = wb.worksheet_range(p_sheet.sheet_for_val) {
            if (p_sheet.validator)(&ws) {
                return Some(Sheet {
                    sheet_name: p_sheet.sheet_name,
                    col_indexers: p_sheet.col_indexers.clone(),
                    col_names: selector.col_names.clone(),
                });
            }
        } else {
            continue;
        }
    }
    None
}

pub fn validate_extractors<RS: Read + Seek>(
    wb: &mut Xlsx<RS>,
    extractors: &[SheetExtractor],
) -> Option<Vec<Sheet>> {
    let mut sheets = vec![];
    for extractor in extractors.iter() {
        let sheet = match extractor {
            SheetExtractor::Multi(extractor) => sheet_from_selector(wb, extractor),
            SheetExtractor::Single(sht) => Some(sht.clone()),
        };
        match sheet {
            Some(s) => sheets.push(s),
            None => return None,
        };
    }
    Some(sheets)
}

pub fn get_botanacor_micro_extractors() -> Vec<SheetExtractor> {
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

    vec![
        test_id_sheet,
        sample_info_sheet,
        tym_values,
        tot_aerobic_values,
        tot_col_values,
    ]
}
