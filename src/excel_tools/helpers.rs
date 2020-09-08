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
            // create a new range and fill with provided default value
            let mut new_range = Range::new((0, 0), (max_row, 0));
            (0..new_range.width()).for_each(|i| new_range.set_value((i as u32, 0), dt.clone()));
            new_range
        },
        ColIndexer::ColFindFunc(func) => func(range, max_row)
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
    ColFindFunc(fn(&Range<DataType>, u32) -> Range<DataType>)
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

pub struct PotentialSheet {
    pub sheet_name: &'static str,
    pub col_indexers: Vec<ColIndexer>,
    pub sheet_for_val: &'static str,
    pub validator: fn(ws: &Range<DataType>) -> bool,
}

pub struct SheetSelector {
    pub col_names: Vec<&'static str>,
    pub potential_sheets: Vec<PotentialSheet>,
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
                    col_names: selector.col_names.clone(),
                    col_indexers: p_sheet.col_indexers.clone(),
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
