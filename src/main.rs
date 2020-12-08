use std::error::Error;
use std::fs;
use std::fs::File;
use std::io;
use std::io::{BufWriter, Write};
use std::path::{Path, PathBuf};

use calamine::{open_workbook, DataType, Reader, Xlsx};
use regex::{Regex, RegexBuilder};

mod excel_tools;
mod extractors;

use extractors::{get_extractors, ExtractionManager};

const DAY_DIR_REGEX: &str = r"\d{2}-[[:alpha:]]{3}-\d{4}";

fn main() -> Result<(), Box<dyn Error>> {
    // setup variables
    let args: Vec<String> = std::env::args().collect();
    println!("Searching in: {}", &args[1]);
    let root_path = Path::new(&args[1]).to_path_buf();
    let company_param = args[2].to_string();
    let test_name_param = args[3].to_string();
    let extraction_t = args[4].to_string();
    let years_regex_str =
        parse_range_arg(args[5].to_string()).expect("Could not parse years parameter");
    let months_regex_str =
        parse_range_arg(args[6].to_string()).expect("Could not parse months parameter");
    let comp_test_string: String = vec![company_param, test_name_param].join("_");
    let extraction_args_string = vec![extraction_t, comp_test_string.clone()].join("_");

    let test_type_regex =
        get_regex(comp_test_string.as_str()).expect("Could not find regex for test type");

    let cg_files = find_cg_files(
        &root_path,
        &test_type_regex,
        years_regex_str,
        months_regex_str,
    );
    println!("Files to be processed: {}", cg_files.len());

    let extraction_manager = get_extractors(&extraction_args_string)?;
    let mut output_data = vec![];
    for file in cg_files.iter() {
        println!("Processing: {}", file.to_str().unwrap());
        let mut rows = file_to_rows_data(file, &extraction_manager)?;
        // Append the file name to the end of each row
        let file_name_dt = DataType::String(String::from(file.to_str().unwrap()));
        rows.iter_mut().for_each(|r| r.push(file_name_dt.clone()));

        output_data.extend(rows);
    }

    let mut header = excel_tools::make_header(&extraction_manager);
    header.push("File Path");
    let output_path = PathBuf::from(extraction_args_string).with_extension("csv");
    let mut output_file = BufWriter::new(File::create(output_path).unwrap());
    write_header(&mut output_file, header)?;
    write_data(&mut output_file, output_data)?;

    println!("\n\nPress ENTER key to exit ...\n");
    io::stdin().read_line(&mut String::new())?;

    Ok(())
}

fn file_to_rows_data(
    file: &PathBuf,
    extraction_manager: &ExtractionManager,
) -> Result<Vec<Vec<DataType>>, Box<dyn Error>> {
    let mut excel: Xlsx<_> = open_workbook(&file)?;
    let sheet_extractors = match extraction_manager {
        ExtractionManager::FileGrain(ex) => ex,
        ExtractionManager::RowGrain(ex) => ex,
    };

    let validated_extractors = excel_tools::validate_extractors(&mut excel, &sheet_extractors);
    if validated_extractors.is_none() {
        println!(
            "File failed sheet validation and will be skipped: {}",
            file.to_str().unwrap()
        );
        return Ok(vec![]);
    }

    let mut active_rows = vec![];
    if let Some(Ok(ws)) = excel.worksheet_range("Master List") {
        match extraction_manager {
            ExtractionManager::FileGrain(_) => active_rows = vec![1], // there will be one row for 'FileGrain' reports
            ExtractionManager::RowGrain(_) => {
                active_rows = excel_tools::find_active_rows(&ws, 1, None);
            }
        }
    }

    let max_row: u32 = 200;
    let mut col_vecs = vec![];
    for sheet in validated_extractors.unwrap().iter() {
        if let Some(Ok(ws)) = excel.worksheet_range(&sheet.sheet_name) {
            col_vecs.extend(excel_tools::extract_sheet_columns(&ws, &sheet, max_row));
        }
    }

    Ok(excel_tools::rows_from_cols(col_vecs, active_rows))
}

fn write_header<W: Write>(dest: &mut W, header: Vec<&str>) -> std::io::Result<()> {
    let len_header = header.len() - 1;
    for (i, h) in header.iter().enumerate() {
        write!(dest, r#""{}""#, h)?;
        if i != len_header {
            write!(dest, ",")?;
        }
    }
    write!(dest, "\r\n")?;

    Ok(())
}

fn write_data<W: Write>(dest: &mut W, data: Vec<Vec<DataType>>) -> std::io::Result<()> {
    for row in data.iter() {
        let row_len = row.len() - 1;
        for (i, d) in row.iter().enumerate() {
            match d {
                DataType::Empty => Ok(()),
                DataType::String(ref s) => write!(dest, r#""{}""#, s),
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

fn find_cg_files(
    root: &PathBuf,
    regex_struct: &TestTypeRegex,
    year_dir_regex: String,
    month_dir_regex: String,
) -> Vec<PathBuf> {
    let year_regex = Regex::new(year_dir_regex.as_str()).unwrap();
    let months_regex = Regex::new(month_dir_regex.as_str()).unwrap();
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

fn parse_range_arg(range_str: String) -> Result<String, Box<dyn Error>> {
    let start_stop: Vec<&str> = range_str.split('-').collect();
    let start: u32 = start_stop[0].parse()?;
    let mut end = start;
    if start_stop.len() == 2 {
        end = start_stop[1].parse()?;
    }
    Ok((start..=end)
        .map(|i| i.to_string())
        .collect::<Vec<String>>()
        .join("|"))
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

fn get_regex(test_type: &str) -> Option<TestTypeRegex> {
    match test_type {
        // Botanacor files
        "botanacor_potency" => Some(TestTypeRegex::new(
            r"^botanacor potency ",
            r"^cert generator botanacor potency .*\.xlsm$",
        )),
        "botanacor_pesticides" => Some(TestTypeRegex::new(
            r"^botanacor pesticides ",
            r"^cert generator botanacor pesticides .*\.xlsm$",
        )),
        "botanacor_metals" => Some(TestTypeRegex::new(
            r"^botanacor metals ",
            r"^cert generator botanacor metals .*\.xlsm$",
        )),
        "botanacor_micro" => Some(TestTypeRegex::new(
            "^validated botanacor micro ",
            r"^cert generator botanacor micro .*\.xlsm$",
        )),

        // Agricor files
        "agricor_micro" => Some(TestTypeRegex::new(
            "^agricor micro ",
            r"^cert generator agricor micro .*\.xlsm$",
        )),
        "agricor_potency" => Some(TestTypeRegex::new(
            "^agricor potency ",
            r"^cert generator agricor potency .*\.xlsm$",
        )),
        _ => None,
    }
}
