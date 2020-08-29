use std::error::Error;
use std::fs;
use std::io;
use std::path::{Path, PathBuf};

use calamine::{open_workbook, Reader, Xlsx};
use regex::{Regex, RegexBuilder};
use std::iter::FlatMap;

const ROOT_DIR: &str = "C:/Users/andre/Desktop/ExtractionTestData";
const YEAR_DIR_REGEX: &str = r"\d{4}";
const MONTH_DIR_REGEX: &str = r"\d{1,2}";
const DAY_DIR_REGEX: &str = r"\d{2}-[[:alpha:]]{3}-\d{4}";

struct TestTypeRegex {
    pub folder: Regex,
    pub file: Regex,
}

impl TestTypeRegex {
    fn new(folder_regex: &str, file_regex: &str) -> Self {
        TestTypeRegex {
            folder: RegexBuilder::new(folder_regex)
                .case_insensitive(true)
                .build()
                .unwrap(),
            file: RegexBuilder::new(file_regex)
                .case_insensitive(true)
                .build()
                .unwrap(),
        }
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    // setup variables
    // TODO: read root dir string from a config file
    let root_path = Path::new(ROOT_DIR).to_path_buf();
    // TODO: get company name & test name from cli
    let company_param = "botanacor".to_string(); // simulating getting from cli
    let test_name_param = "micro".to_string(); // simulating getting from cli
    let test_type: String = vec![company_param, test_name_param].join("_");

    // let regex_map = build_regex_map();
    // let test_type_regex = regex_map.get(test_type.as_str()).unwrap();
    let test_type_regex =
        get_regex(test_type.as_str()).expect("Could not find regex for test type");

    // let cg_files = find_cg_files(&root_path, &test_type_regex);
    let cg_files = get_cg_files(&root_path, &test_type_regex)?;
    for file in cg_files {
        try_open_wb(&file);
    }

    Ok(())
}

fn find_cg_files(root: &PathBuf, regex_struct: &TestTypeRegex) -> Vec<PathBuf> {
    let year_regex = Regex::new(YEAR_DIR_REGEX).unwrap();
    let months_regex = Regex::new(MONTH_DIR_REGEX).unwrap();
    let days_regex = RegexBuilder::new(DAY_DIR_REGEX)
        .case_insensitive(true)
        .build()
        .unwrap();

    match_child_paths(&root, &year_regex)
        .iter()
        .filter(|p| p.is_dir())
        .flat_map(|p| match_child_paths(&p, &months_regex))
        .filter(|p| p.is_dir())
        .flat_map(|p| match_child_paths(&p, &days_regex))
        .filter(|p| p.is_dir())
        .flat_map(|p| match_child_paths(&p, &regex_struct.folder))
        .filter(|p| p.is_dir())
        .flat_map(|p| match_child_paths(&p, &regex_struct.file))
        .collect()
}

fn get_cg_files(root: &PathBuf, regex_struct: &TestTypeRegex) -> Result<Vec<PathBuf>, io::Error> {
    let year_regex = Regex::new(YEAR_DIR_REGEX).unwrap();
    let months_regex = Regex::new(MONTH_DIR_REGEX).unwrap();
    let days_regex = RegexBuilder::new(DAY_DIR_REGEX)
        .case_insensitive(true)
        .build()
        .unwrap();

    let year_dirs: Vec<PathBuf> = fs::read_dir(root)?
        .filter_map(|r| filter_by_filename(r, &year_regex))
        .map(|dir_ent| dir_ent.path())
        .filter(|p| p.is_dir())
        .collect();

    let mut month_dirs: Vec<PathBuf> = vec!();
    for dir in year_dirs {
        let dir_reader = fs::read_dir(dir)?;
        let dirs = dir_reader
            .filter(|dr| dr.is_ok())
            .filter_map(|dr| filter_file(dr, &months_regex))
            .filter(|d| d.is_dir());
        month_dirs.extend(dirs)
    }

    let day_dirs: Vec<PathBuf> = month_dirs
        .iter()
        .map(fs::read_dir)
        .filter(|d| d.is_ok())
        .map(|d| d.unwrap())
        .flat_map(|read_dir| read_dir.filter_map(|d| filter_file(d, &days_regex)))
        .filter(|p| p.is_dir())
        .collect();

    let cg_dirs: Vec<PathBuf> = day_dirs
        .iter()
        .map(fs::read_dir)
        .filter(|d| d.is_ok())
        .map(|d| d.unwrap())
        .flat_map(|read_dir| read_dir.filter_map(|d| filter_file(d, &regex_struct.folder)))
        .filter(|p| p.is_dir())
        .collect();

    let files: Vec<PathBuf> = cg_dirs
        .iter()
        .map(fs::read_dir)
        .filter(|d| d.is_ok())
        .map(|d| d.unwrap())
        .flat_map(|read_dir| read_dir.filter_map(|d| filter_file(d, &regex_struct.file)))
        .collect();

    Ok(files)
}

fn filter_file(dir: Result<fs::DirEntry, io::Error>, pattern: &Regex) -> Option<PathBuf> {
    match dir {
        Ok(dir_entry) => {
            if pattern.is_match(dir_entry.path().file_name().unwrap().to_str().unwrap()) {
                Some(dir_entry.path())
            }
            else {
                None
            }
        },
        Err(_) => None
    }
}

fn match_child_paths(parent_dir: &PathBuf, child_regex: &Regex) -> Vec<PathBuf> {
    fs::read_dir(parent_dir)
        .unwrap()
        .filter_map(|r| filter_by_filename(r, child_regex))
        .map(|e| e.path())
        .collect()
}

fn filter_by_filename(
    dir: Result<fs::DirEntry, io::Error>,
    pattern: &Regex,
) -> Option<fs::DirEntry> {
    match dir {
        Ok(entry) => {
            if pattern.is_match(entry.path().file_name().unwrap().to_str().unwrap()) {
                Some(entry)
            } else {
                None
            }
        }
        Err(_) => None,
    }
}

fn try_open_wb(path: &PathBuf) {
    println!("Opening {}", path.display());
    let mut excel: Xlsx<_> = open_workbook(path).unwrap();
    if let Some(Ok(r)) = excel.worksheet_range("Master List") {
        let b1 = r.get_value((0, 1));
        println!("{:?}", b1.unwrap());
        // let header = r.rows().next().unwrap();
        // println!("row[0]={:?}", header[1]);
    }
}

fn get_regex(test_type: &str) -> Option<TestTypeRegex> {
    match test_type {
        "botanacor_potency" => Some(TestTypeRegex::new(
            r"^botanacor potency ",
            r"^cert generator botanacor potency .*\.xlsm$",
        )),
        "botanacor_metals" => Some(TestTypeRegex::new(
            r"^botanacor metals ",
            r"^cert generator botanacor metals .*\.xlsm$",
        )),
        "botanacor_micro" => Some(TestTypeRegex::new(
            "^validated botanacor micro ",
            r"^cert generator botanacor micro .*\.xlsm$",
        )),
        _ => None,
    }
}
