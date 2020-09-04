use calamine::{DataType, Range};

pub fn header_match(
    ws: &Range<DataType>,
    starts_with: &str,
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
                // println!("header text: {}", s);
                // println!("matching: {}", starts_with);
                if s.starts_with(starts_with) {
                    return ws.range(
                        (range_offset, col as u32),
                        (range_offset + row_count, col as u32),
                    );
                }
            }
            _ => continue,
        }
    }
    // return range filled with a default string
    let mut default = Range::new((0, 0), (row_count, 0));
    (0..default.height())
        .for_each(|i| default.set_value((0, i as u32), DataType::String("NA".to_string())));
    default
}
