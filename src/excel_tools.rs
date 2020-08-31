use calamine::{DataType, Range};

type Accessor = fn(&Range<DataType>) -> bool;

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
                        continue
                    }
                    else {
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

pub fn row_values_from_sheet<'a>(
    range: &'a Range<DataType>,
    sheet: &Sheet,
    row: u32,
) -> Vec<DataType> {
    let mut values = vec![];
    for col in sheet.col_numbers.iter() {
        values.push(range.get_value((row, *col)).unwrap().clone());
    }

    values
}

pub struct Sheet {
    pub sheet_name: &'static str,
    pub col_names: Vec<&'static str>,
    pub col_numbers: Vec<u32>,
    validator: Option<Accessor>,
}

pub fn get_botanacor_micro_sheets() -> Vec<Sheet> {
    let sample_info = Sheet {
        sheet_name: "Master List",
        col_names: vec!["Test Id", "Sample Name"],
        col_numbers: vec![1, 3],
        validator: Some(|_ws: &Range<DataType>| true),
    };

    vec![sample_info]
}
