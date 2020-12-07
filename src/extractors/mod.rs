mod extraction_manager;
mod results;
mod controls;

pub use extraction_manager::ExtractionManager;

pub fn get_extractors(extractors_name: &str) -> Result<ExtractionManager, String> {
    match extractors_name {
        // * Results extraction
        // Botanacor
        "results_botanacor_micro" => Ok(results::botanacor_micro::get_extractors()),
        "results_botanacor_pesticides" => Ok(results::botanacor_pesticides::get_extractors()),
        // Agricor
        "results_agricor_micro" => Ok(results::agricor_micro::get_extractors()),
        "results_agricor_potency" => Ok(results::agricor_potency::get_extractors()),
        // * Control charts extraction
        // Agricor
        "controls_agricor_potency" => Ok(controls::agricor_potency::get_extractors()),
        // No match
        _ => Err(format!(
            "Could not find data extraction configuration for {}",
            extractors_name
        )),
    }
}
