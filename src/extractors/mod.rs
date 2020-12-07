mod botanacor_micro;
mod botanacor_pesticides;
mod agricor_micro;
mod agricor_potency;

use crate::excel_tools::SheetExtractor;

pub fn get_extractors(extractors_name: &str) -> Result<Vec<SheetExtractor>, String> {
    match extractors_name {
        // Botanacor files
        "botanacor_micro" => Ok(botanacor_micro::get_extractors()),
        "botanacor_pesticides" => Ok(botanacor_pesticides::get_extractors()),

        // Agricor files
        "agricor_micro" => Ok(agricor_micro::get_extractors()),
        "agricor_potency" => Ok(agricor_potency::get_extractors()),
        // No match
        _ => Err(format!(
            "Could not find data extraction configuration for {}",
            extractors_name
        )),
    }
}
