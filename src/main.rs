use anyhow::Result;
use calamine::{open_workbook, Xlsx, RangeDeserializerBuilder, Reader};
use native_dialog::FileDialog;
use serde::{Deserialize, Serialize};

#[serde(rename_all="PascalCase")]
#[derive(Debug, Deserialize, Serialize)]
struct Row {
    pub term: String,
    pub translation: String,
    pub language: String,
}

fn main() -> Result<()> {
    println!("Hello, world!");
    let result = FileDialog::new()
        .set_location("~")
        .add_filter("text", &["xlsx"])
        .show_open_single_file()?;

    if let Some(excel) = result {
        println!("{:?}", excel);
        let mut workbook: Xlsx<_> = open_workbook(excel)?;
        let range = workbook
            .worksheet_range("Sheet 1").expect("Unable to load Sheet 1").unwrap();

        let  iter = RangeDeserializerBuilder::new().from_range::<_, Row>(&range)?;
        for ele in iter {
            println!("{:?}", ele);
        }
    };

    std::process::abort();
}
