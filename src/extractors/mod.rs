mod botanacor_micro;
mod agricor_micro;

use crate::excel_tools::SheetExtractor;

pub fn get_extractors(extractors_name: &str) -> Result<Vec<SheetExtractor>, String> {
    match extractors_name {
        "botanacor_micro" => Ok(botanacor_micro::get_extractors()),
        "agricor_micro" => Ok(agricor_micro::get_extractors()),
        _ => Err(format!(
            "Could not find data extraction configuration for {}",
            extractors_name
        )),
    }
}
