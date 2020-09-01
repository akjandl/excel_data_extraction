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

pub fn extract_column(range: &Range<DataType>, col_num: u32, max_row: u32) -> Range<DataType> {
    range.range((1, col_num), (max_row, col_num))
}

pub fn extract_sheet_columns(
    range: &Range<DataType>,
    sheet: &Sheet,
    max_row: u32,
) -> Vec<Range<DataType>> {
    let mut cols = vec![];
    for col in sheet.col_indexers.iter() {
        cols.push(extract_column(range, *col, max_row))
    }
    cols
}

pub fn rows_from_cols(cols: Vec<Range<DataType>>, active_rows: Vec<u32>) -> Vec<Vec<DataType>> {
    let mut rows_data = vec![];
    for row in active_rows.iter() {
        let mut cur_row = vec![];
        for col in cols.iter() {
            cur_row.push(
                col.get_value((*row, col.start().unwrap().1))
                    .unwrap()
                    .clone(),
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
pub struct Sheet {
    pub sheet_name: &'static str,
    pub col_names: Vec<&'static str>,
    pub col_indexers: Vec<u32>,
}

pub enum SheetExtractor {
    Single(Sheet),
    Multi(SheetSelector),
}

struct PotentialSheet {
    sheet_name: &'static str,
    col_indexers: Vec<u32>,
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
    let test_id_sheet = SheetExtractor::Multi(SheetSelector {
        col_names: vec!["Test Id"],
        potential_sheets: vec![
            PotentialSheet {
                sheet_name: "AgrBotMap",
                col_indexers: vec![3],
                sheet_for_val: "AgrBotMap",
                validator: |ws| {
                    if let Some(dt) = ws.get_value((0, 3)) {
                        match dt {
                            DataType::String(s) => s == &"AgricorSampleName".to_string(),
                            _ => false,
                        }
                    } else {
                        false
                    }
                },
            },
            PotentialSheet {
                sheet_name: "Master List",
                col_indexers: vec![1],
                sheet_for_val: "Master List",
                validator: |ws| {
                    if let Some(dt) = ws.get_value((0, 1)) {
                        match dt {
                            DataType::String(s) => s == &"Test Id".to_string(),
                            _ => false,
                        }
                    } else {
                        false
                    }
                },
            },
        ],
    });

    let tym_values = SheetExtractor::Single(Sheet {
        sheet_name: "TYM Values",
        col_names: vec!["TYM CFU Count", "TYM Dil. Plate", "TYM Reported CFU/g"],
        col_indexers: vec![4, 5, 6],
    });

    vec![test_id_sheet, tym_values]
}
