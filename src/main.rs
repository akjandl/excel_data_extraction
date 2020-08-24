use std::error::Error;
use std::io;
use std::fs;
use std::path::{PathBuf, Path};
use regex::{Regex, RegexBuilder};
use std::collections::HashMap;


const ROOT_DIR: &str = "C:/Users/andre/Desktop/ExtractionTestData";

struct TestTypeRegex {
    folder: Regex,
    file: Regex,
}

fn main() -> Result<(), Box<dyn Error>> {
    // setup variables
    // TODO: get company name & test name from cli
    let root_path = Path::new(ROOT_DIR).to_path_buf();
    let company_param = "botanacor".to_string(); // simulating getting from cli
    let test_name_param = "micro".to_string(); // simulating getting from cli
    let test_type = vec!(company_param, test_name_param).join("_");

    // Cert gen directory regex lookup HashMap
    let mut dirs_regex: HashMap<&str, &str> = HashMap::new();
    dirs_regex.insert("botanacor_potency", r"^botanacor potency ");
    dirs_regex.insert("botanacor_metals", r"^botanacor metals ");
    dirs_regex.insert("botanacor_micro", r"^validated botanacor micro ");

    // Cert gen filename regex lookup HashMap
    let mut files_regex: HashMap<&str, &str> = HashMap::new();
    files_regex.insert("botanacor_potency", r"^cert generator botanacor potency .*\.xlsm$");
    files_regex.insert("botanacor_metals", r"^cert generator botanacor metals .*\.xlsm$");
    files_regex.insert("botanacor_micro", r"^cert generator botanacor micro .*\.xlsm$");

    let dir_regex = dirs_regex.get(test_type.as_str()).expect("Regex for test type directories not found");
    let filename_regex = files_regex.get(test_type.as_str()).expect("Regex for test type files not found");
    let bot_potency = TestTypeRegex {
        folder: RegexBuilder::new(dir_regex).case_insensitive(true).build().unwrap(),
        file: RegexBuilder::new(filename_regex).case_insensitive(true).build().unwrap(),
    };

    let bot_pot_files = find_cg_files(&root_path, &bot_potency);
    bot_pot_files.into_iter().map(|e| println!("{}", e.display())).for_each(drop);

    Ok(())
}

fn find_cg_files(root: &PathBuf, regex_struct: &TestTypeRegex) -> Vec<PathBuf> {
    let year_regex = Regex::new(r"\d{4}").unwrap();
    let months_regex = Regex::new(r"\d{1,2}").unwrap();
    let dates_regex = RegexBuilder::new(r"\d{2}-[[:alpha:]]{3}-\d{4}").case_insensitive(true).build().unwrap();
    match_child_paths(&root, &year_regex)
        .iter().filter(|e| e.is_dir())
        .flat_map(|p| match_child_paths(&p, &months_regex)).filter(|p| p.is_dir())
        .flat_map(|p| match_child_paths(&p, &dates_regex)).filter(|p| p.is_dir())
        .flat_map(|p| match_child_paths(&p, &regex_struct.folder)).filter(|p| p.is_dir())
        .flat_map(|p| match_child_paths(&p, &regex_struct.file)).collect()
}

fn match_child_paths(parent_dir: &PathBuf, child_regex: &Regex) -> Vec<PathBuf> {
    fs::read_dir(parent_dir).unwrap()
        .filter_map(|r| filter_by_filename(r, child_regex))
        .map(|e| e.path())
        .collect()
}

fn filter_by_filename(dir: Result<fs::DirEntry, io::Error>, pattern: &Regex) -> Option<fs::DirEntry> {
    match dir {
        Ok(entry) => if pattern.is_match(entry.path().file_name().unwrap().to_str().unwrap()) {
            Some(entry)
        } else {
            None
        },
        Err(_) => None
    }
}
