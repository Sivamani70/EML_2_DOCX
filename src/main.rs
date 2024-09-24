mod mail;
mod newdoc;

use clap::{arg, Parser};
use mail::Mail;
use newdoc::NewDocx;
use std::path::PathBuf;

#[derive(Parser, Debug)]
#[command(
    version = "1.0.0",
    about = "The application is designed to parse Outlook (.eml) files and generate structured Word document based on the extracted email headers."
)]
struct Args {
    #[arg(
        short = 'i',
        long = "in-file",
        value_name = "FILE PATH",
        help = "Input file"
    )]
    in_file: String,

    #[arg(
        short = 'o',
        long = "out-file",
        value_name = "FILE PATH",
        help = "Output file"
    )]
    out_file: String,

    #[arg(
        short = 'n',
        long = "i-num",
        value_name = "Incident NUMBER",
        help = "Integer value"
    )]
    i_num: String,
}

fn main() {
    let args = Args::parse();

    let in_file = args.in_file;
    let out_file = args.out_file;
    let incident_number = args.i_num;
    let eml = Mail::new(PathBuf::from(in_file));
    let data = eml.get_content();
    let (headers, body_headers, _content) = eml.parse(data.as_bytes());

    let new_docx = NewDocx::new(PathBuf::from(out_file), incident_number);
    let doc = new_docx.generate_content(headers, body_headers);
    new_docx.create_docx(doc);
}
