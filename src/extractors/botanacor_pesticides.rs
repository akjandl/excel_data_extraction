use calamine::{DataType, Range};
use std::rc::Rc;

use crate::excel_tools::column_finders::header_match;
use crate::excel_tools::{ColIndexer, PotentialSheet, Sheet, SheetExtractor, SheetSelector};

fn find_col(starts_with: &'static str) -> Rc<dyn Fn(&Range<DataType>, u32) -> Range<DataType>> {
    Rc::new(move |ws: &Range<DataType>, row_count: u32| {
        header_match(ws, starts_with, 0, row_count, None)
    })
}

pub fn get_extractors() -> Vec<SheetExtractor> {
    let early_2019_validator = |ws: &Range<DataType>| -> bool {
        if let Some(dt) = ws.get_value((0, 2)) {
            match dt {
                DataType::String(s) => s.starts_with("Customer Name"),
                _ => false,
            }
        } else {
            false
        }
    };
    let with_test_id_validator = |ws: &Range<DataType>| -> bool {
        if let Some(dt) = ws.get_value((0, 1)) {
            match dt {
                DataType::String(s) => s.starts_with("Test Id"),
                _ => false,
            }
        } else {
            false
        }
    };

    let master_list = SheetExtractor::Multi(SheetSelector {
        col_names: vec!["Test Id", "Customer Name", "Sample Name", "Report Type"],
        potential_sheets: vec![
            PotentialSheet {
                sheet_name: "Master List",
                sheet_for_val: "Master List",
                validator: with_test_id_validator,
                col_indexers: vec![
                    ColIndexer::ColFindFunc(find_col("Test Id")),
                    ColIndexer::ColFindFunc(find_col("Testing Company Name")),
                    ColIndexer::ColFindFunc(find_col("Sample Info")),
                    ColIndexer::ColFindFunc(find_col("Report Type")),
                ],
            },
            PotentialSheet {
                sheet_name: "Master List",
                sheet_for_val: "Master List",
                validator: early_2019_validator,
                col_indexers: vec![
                    ColIndexer::DefaultValue(DataType::String("NA".to_string())),
                    ColIndexer::ColFindFunc(find_col("Customer Name")),
                    ColIndexer::ColFindFunc(find_col("Sample Info")),
                    ColIndexer::ColFindFunc(find_col("Report Type")),
                ],
            },
        ],
    });

    let results = SheetExtractor::Single(Sheet {
        sheet_name: "Sample Data",
        col_names: vec![
            "Unit",
            "Acephate",
            "Oxamyl",
            "Methomyl",
            "Flonicamid",
            "Thiamethoxam",
            "Dimethoate",
            "Imidacloprid",
            "Acetamiprid",
            "Thiacloprid",
            "Dichlorvos",
            "Propoxur",
            "Carbofuran",
            "Carbaryl",
            "Imazalil",
            "Metalaxyl",
            "Naled",
            "Spiroxamine 1",
            "Spiroxamine 2",
            "Methiocarb",
            "Chlorantraniliprole",
            "Fludioxonil",
            "Paclobutrazol",
            "Prophos",
            "Boscalid",
            "Myclobutanil",
            "Phosmet",
            "Malathion",
            "Azoxystrobin",
            "Bifenazate",
            "Spirotetramat",
            "Fipronil",
            "Tebuconazole",
            "Fenoxycarb",
            "Diazinon",
            "Kresoxim-methyl",
            "MGK 264 1",
            "Clofentezine",
            "MGK 264 2",
            "Trifloxystrobin",
            "Spinosad A",
            "Spiromesifen",
            "Spinosad D",
            "Etoxazole",
            "Chlorpyrifos",
            "Hexythiazox",
            "E-Fenpyroximate",
            "Pyridaben",
            "Avermectin",
            "Permethrin",
            "Etofenprox",
        ],
        col_indexers: vec![
            ColIndexer::ColFindFunc(find_col("Units")),
            ColIndexer::ColFindFunc(find_col("Acephate")),
            ColIndexer::ColFindFunc(find_col("Oxamyl")),
            ColIndexer::ColFindFunc(find_col("Methomyl")),
            ColIndexer::ColFindFunc(find_col("Flonicamid")),
            ColIndexer::ColFindFunc(find_col("Thiamethoxam")),
            ColIndexer::ColFindFunc(find_col("Dimethoate")),
            ColIndexer::ColFindFunc(find_col("Imidacloprid")),
            ColIndexer::ColFindFunc(find_col("Acetamiprid")),
            ColIndexer::ColFindFunc(find_col("Thiacloprid")),
            ColIndexer::ColFindFunc(find_col("Dichlorvos")),
            ColIndexer::ColFindFunc(find_col("Propoxur")),
            ColIndexer::ColFindFunc(find_col("Carbofuran")),
            ColIndexer::ColFindFunc(find_col("Carbaryl")),
            ColIndexer::ColFindFunc(find_col("Imazalil")),
            ColIndexer::ColFindFunc(find_col("Metalaxyl")),
            ColIndexer::ColFindFunc(find_col("Naled")),
            ColIndexer::ColFindFunc(find_col("Spiroxamine 1")),
            ColIndexer::ColFindFunc(find_col("Spiroxamine 2")),
            ColIndexer::ColFindFunc(find_col("Methiocarb")),
            ColIndexer::ColFindFunc(find_col("Chlorantraniliprole")),
            ColIndexer::ColFindFunc(find_col("Fludioxonil")),
            ColIndexer::ColFindFunc(find_col("Paclobutrazol")),
            ColIndexer::ColFindFunc(find_col("Prophos")),
            ColIndexer::ColFindFunc(find_col("Boscalid")),
            ColIndexer::ColFindFunc(find_col("Myclobutanil")),
            ColIndexer::ColFindFunc(find_col("Phosmet")),
            ColIndexer::ColFindFunc(find_col("Malathion")),
            ColIndexer::ColFindFunc(find_col("Azoxystrobin")),
            ColIndexer::ColFindFunc(find_col("Bifenazate")),
            ColIndexer::ColFindFunc(find_col("Spirotetramat")),
            ColIndexer::ColFindFunc(find_col("Fipronil")),
            ColIndexer::ColFindFunc(find_col("Tebuconazole")),
            ColIndexer::ColFindFunc(find_col("Fenoxycarb")),
            ColIndexer::ColFindFunc(find_col("Diazinon")),
            ColIndexer::ColFindFunc(find_col("Kresoxim-methyl")),
            ColIndexer::ColFindFunc(find_col("MGK 264 1")),
            ColIndexer::ColFindFunc(find_col("Clofentezine")),
            ColIndexer::ColFindFunc(find_col("MGK 264 2")),
            ColIndexer::ColFindFunc(find_col("Trifloxystrobin")),
            ColIndexer::ColFindFunc(find_col("Spinosad A")),
            ColIndexer::ColFindFunc(find_col("Spiromesifen")),
            ColIndexer::ColFindFunc(find_col("Spinosad D")),
            ColIndexer::ColFindFunc(find_col("Etoxazole")),
            ColIndexer::ColFindFunc(find_col("Chlorpyrifos")),
            ColIndexer::ColFindFunc(find_col("Hexythiazox")),
            ColIndexer::ColFindFunc(find_col("E-Fenpyroximate")),
            ColIndexer::ColFindFunc(find_col("Pyridaben")),
            ColIndexer::ColFindFunc(find_col("Avermectin")),
            ColIndexer::ColFindFunc(find_col("Permethrin")),
            ColIndexer::ColFindFunc(find_col("Etofenprox")),
        ],
    });

    vec![master_list, results]
}
