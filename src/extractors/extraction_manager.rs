use crate::excel_tools::SheetExtractor;

pub enum ExtractionManager {
    FileGrain(Vec<SheetExtractor>),
    RowGrain(Vec<SheetExtractor>),
}
