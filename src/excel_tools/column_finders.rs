use calamine::{DataType, Range};

pub enum MatchMethod {
    StartsWith(&'static str),
    Exact(&'static str),
}

pub fn header_match(
    ws: &Range<DataType>,
    match_method: MatchMethod,
    header_row: u32,
    row_count: u32,
    start_offset: Option<u32>,
) -> Range<DataType> {
    // set row offset for the returned range
    let range_offset: u32;
    match start_offset {
        Some(offset) => range_offset = offset,
        None => range_offset = 1,
    };

    let header_max_col = (ws.width() - 1) as u32;
    let header = ws.range((header_row, 0), (header_row, header_max_col));
    for cell in header.cells() {
        let (_row, col, val) = cell;
        match val {
            DataType::String(s) => {
                match match_method {
                    MatchMethod::StartsWith(match_str) => {
                        if s.starts_with(match_str) {
                            return ws.range(
                                (range_offset, col as u32),
                                (range_offset + row_count, col as u32),
                            );
                        }
                    }
                    MatchMethod::Exact(match_str) => {
                        if s == match_str {
                            return ws.range(
                                (range_offset, col as u32),
                                (range_offset + row_count, col as u32),
                            )
                        }
                    }
                }
            }
            _ => continue,
        }
    }
    // return range filled with a default string
    let mut default = Range::new((0, 0), (row_count, 0));
    (0..default.height())
        .for_each(|i| default.set_value((i as u32, 0), DataType::String("NA".to_string())));
    default
}
