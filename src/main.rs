use std::error::Error;
use std::fs;
use std::fs::File;
use std::io;
use std::io::{BufWriter, Write};
use std::path::{Path, PathBuf};

use calamine::{open_workbook, DataType, Reader, Xlsx};
use regex::{Regex, RegexBuilder};

mod excel_tools;

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

    let test_type_regex =
        get_regex(test_type.as_str()).expect("Could not find regex for test type");

    let cg_files = find_cg_files(&root_path, &test_type_regex);
    let sheets = excel_tools::get_botanacor_micro_sheets();

    let mut output_data = vec![];
    for file in cg_files {
        println!(
            "Processing: {}",
            file.file_name().unwrap().to_str().unwrap()
        );
        let mut excel: Xlsx<_> = open_workbook(&file)?;
        let mut active_rows = vec![];
        if let Some(Ok(ws)) = excel.worksheet_range("Master List") {
            active_rows = excel_tools::find_active_rows(&ws, 1, None);
        }
        for row in active_rows.iter() {
            let mut row_data = vec![];
            for sheet in sheets.iter() {
                if let Some(Ok(ws)) = excel.worksheet_range(&sheet.sheet_name) {
                    let row_vals = excel_tools::row_values_from_sheet(&ws, sheet, *row);
                    row_data.extend(row_vals);
                }
            }
            println!("{:?}", row_data);
            output_data.push(row_data);
        }
    }

    let output_path = PathBuf::from("output_file").with_extension("csv");
    let mut output_file = BufWriter::new(File::create(output_path).unwrap());
    write_data(&mut output_file, output_data)?;

    Ok(())
}

fn write_data<W: Write>(dest: &mut W, data: Vec<Vec<DataType>>) -> std::io::Result<()> {
    for row in data.iter() {
        let row_len = row.len() - 1;
        for (i, d) in row.iter().enumerate() {
            match d {
                DataType::Empty => Ok(()),
                DataType::String(ref s) => write!(dest, "{}", s),
                DataType::Float(ref f) => write!(dest, "{}", f),
                DataType::Int(ref i) => write!(dest, "{}", i),
                DataType::Error(ref e) => write!(dest, "{:?}", e),
                DataType::Bool(ref b) => write!(dest, "{}", b),
            }?;
            if i != row_len {
                write!(dest, ",")?;
            }
        }
        write!(dest, "\r\n")?;
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
