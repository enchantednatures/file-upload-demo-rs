use std::path::PathBuf;

use anyhow::Result;
use calamine::{open_workbook, RangeDeserializerBuilder, Reader, Xlsx};
use clap::Parser;
use native_dialog::FileDialog;
use serde::{Deserialize, Serialize};

#[derive(Debug, Deserialize, Serialize)]
#[serde(rename_all = "PascalCase")]
struct Row {
    pub term: String,
    pub translation: String,
    pub language: String,
}

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct CliArgs {
    #[arg(short, long)]
    path: Option<String>,
}

fn main() -> Result<()> {
    println!("Hello, world!");

    let args = CliArgs::parse();
    let result: Option<PathBuf>;

    if let Some(arg_file) = args.path {
        result = Some(PathBuf::from(arg_file));
    } else {
        result = FileDialog::new()
            .set_location("~")
            .add_filter("text", &["xlsx"])
            .show_open_single_file()?;
    }

    if let Some(excel) = result {
        println!("{:?}", excel);
        let mut workbook: Xlsx<_> = open_workbook(excel)?;
        let range = workbook
            .worksheet_range("Sheet 1")
            .expect("Unable to load Sheet 1")
            .unwrap();

        let iter = RangeDeserializerBuilder::new().from_range::<_, Row>(&range)?;
        for ele in iter {
            println!("{:?}", ele);
        }
    };

    std::process::exit(0);
}
