use docx_rs::{Docx, PageMargin, Paragraph, Run, RunFonts};
use mailparse::parse_mail;
use std::{
    collections::HashMap,
    env,
    fs::{self},
    path::{Path, PathBuf},
};

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.len() < 4 {
        panic!("Provide input file, outfile and Incident number");
    }

    let in_file = args[1].clone();
    let out_file = args[2].clone();
    let incident_number = args[3].clone();
    let eml = Mail::new(PathBuf::from(in_file));
    let data = eml.get_content();
    let (headers, body_headers, content) = eml.parse(data.as_bytes());

    let new_docx = NewDocx::new(PathBuf::from(out_file), incident_number);
    let doc = new_docx.generate_content(headers, body_headers);
    new_docx.create_docx(doc);
}

struct Mail {
    eml_path: PathBuf,
}

impl Mail {
    pub fn new(input_file: PathBuf) -> Self {
        Self {
            eml_path: input_file,
        }
    }

    pub fn get_content(&self) -> String {
        let content = fs::read_to_string(&self.eml_path);

        match content {
            Ok(data) => return data,
            Err(err) => {
                eprintln!("Unable to read data from the eml file:\n{err}");
                panic!("{err}")
            }
        }
    }

    pub fn parse(
        &self,
        data: &[u8],
    ) -> (
        HashMap<String, String>,
        Vec<HashMap<String, String>>,
        Vec<String>,
    ) {
        let parsed_mail = parse_mail(data);
        let mut main_headers = HashMap::<String, String>::new();
        let mut body_headers_list = Vec::<HashMap<String, String>>::new();
        let mut body_content = Vec::<String>::new();

        let (headers, sub_parts, _ctype) = match parsed_mail {
            Ok(p) => (p.headers, p.subparts, p.ctype.params),
            Err(err) => {
                eprintln!("Unable to Parse the eml file:\n{err}");
                panic!("{err}")
            }
        };

        for h in headers {
            let key = h.get_key();
            let value = h.get_value();
            Self::add_to_map(&mut main_headers, key, value);
        }

        for sp in sub_parts.iter() {
            let headers = sp.get_headers();
            let body = sp.get_body();

            match body {
                Ok(data) => body_content.push(data),
                Err(err) => {
                    eprintln!("Error while parsing the body {err}");
                }
            }

            for h in headers {
                let key = h.get_key();
                let value = h.get_value();
                let mut body_headers = HashMap::<String, String>::new();
                Self::add_to_map(&mut body_headers, key, value);
                body_headers_list.push(body_headers);
            }
        }
        return (main_headers, body_headers_list, body_content);
    }

    fn add_to_map(h_map: &mut HashMap<String, String>, key: String, value: String) {
        let pos = h_map.insert(key.to_owned(), value.to_owned());

        match pos {
            Some(old_value) => {
                println!("The key: {key} exists in the map.\nOld value: {old_value} is replaced by \n->\t New Value:{value}");
            }
            None => {
                println!("New value inserted with Key: {key}");
            }
        }
    }
}

struct NewDocx {
    docx_path: PathBuf,
    i_number: String,
}

impl NewDocx {
    pub fn new(path: PathBuf, incident_number: String) -> Self {
        Self {
            docx_path: path,
            i_number: incident_number,
        }
    }

    pub fn create_docx(&self, doc: Docx) {
        let path = Path::new(&self.docx_path);
        let file = fs::File::create(path);

        let file = match file {
            Ok(f) => f,
            Err(err) => {
                eprintln!("Unable to create word document");
                panic!("{err}");
            }
        };

        match doc.build().pack(file) {
            Ok(_) => {
                println!("Document creation completed")
            }
            Err(err) => {
                eprintln!("{err}");
                panic!("{err}")
            }
        }
    }

    pub fn generate_content(
        &self,
        headers: HashMap<String, String>,
        b_headers: Vec<HashMap<String, String>>,
    ) -> Docx {
        let heading = format!("INCIDENT {}", &self.i_number);
        let se_line = format!("Please find the following initial analysis details.");
        let date = format!("Date:\t {}", Self::get_values("Date", &headers));
        let subject = format!("Subject:\t {}", Self::get_values("Subject", &headers));
        let from = format!("Sender Id:\t {}", Self::get_values("From", &headers));
        let to = format!("Recipient Id:\t {}", Self::get_values("To", &headers));
        let return_path = format!(
            "Return Path:\t {}",
            Self::get_values("Return-Path", &headers)
        );
        let h_content_type = format!("Subject:\t {}", Self::get_values("Content-Type", &headers));
        let spf = format!("Subject:\t {}", Self::get_values("Received-Spf", &headers));
        let auth_result = format!(
            "Subject:\t {}",
            Self::get_values("Authentication-Results", &headers)
        );
        let form_address = Self::get_values("From", &headers);
        let parts: Vec<&str> = form_address.trim().split("@").collect();
        let domain = if parts.len() != 2 {
            "Not able to extract domain".to_string()
        } else {
            match parts.last() {
                Some(d) => d.replace(">", ""),
                None => "Not able to extract domain".to_string(),
            }
        };

        let s_domain = format!("Domain:\t {}", domain);

        let mut have_attachments = false;
        let mut count = 0;
        for bh in b_headers {
            let att = Self::get_values("Content-Disposition", &bh);
            if att != "NA" {
                have_attachments = true;
                count += 1;
            }
        }

        let attachments = if have_attachments {
            format!("Attachments:\t Yes")
        } else {
            format!("Attachments:\t No")
        };

        let analysis_pt = format!("User received a mail from {} which was detected as a ****-suspicious mail. As per the initial analysis we gathered that the mail came from {}.", Self::get_values("From", &headers), domain);

        let anlysis_url_atch = format!(
            "We also observed that there are ***** URL(s) and {} Attachment(s) in this email body.",
            count
        );

        let mut docx = Docx::new();

        docx = docx.add_paragraph(Self::head(&heading, "Blue", 36));
        docx = docx.add_paragraph(Self::build_paragraph(&se_line, "Black", 22));
        docx = docx.add_paragraph(Self::build_paragraph(&date, "Black", 22));
        docx = docx.add_paragraph(Self::build_paragraph(&subject, "Black", 22));
        docx = docx.add_paragraph(Self::build_paragraph(&from, "Black", 22));
        docx = docx.add_paragraph(Self::build_paragraph(&to, "Black", 22));
        docx = docx.add_paragraph(Self::build_paragraph(&s_domain, "Black", 22));
        docx = docx.add_paragraph(Self::build_paragraph(&attachments, "Black", 22));
        docx = docx.add_paragraph(Self::build_paragraph(
            &"Attachments (Malicious)(Y/N):\t N/A".to_string(),
            "Black",
            22,
        ));
        docx = docx.add_paragraph(Self::build_paragraph(
            &"URL:\t N/A".to_string(),
            "Black",
            22,
        ));
        docx = docx.add_paragraph(Self::build_paragraph(
            &"URL (Malicious)(Y/N):\t N/A".to_string(),
            "Black",
            22,
        ));
        docx = docx.add_paragraph(Paragraph::new());
        docx = docx.add_paragraph(Self::build_paragraph(
            &"Domain Analysis".to_string(),
            "Blue",
            28,
        ));
        docx = docx.add_paragraph(Paragraph::new());
        docx = docx.add_paragraph(Self::build_paragraph(&"Analysis".to_string(), "Blue", 28));
        docx = docx.add_paragraph(Paragraph::new());
        docx = docx.add_paragraph(Self::build_paragraph(&analysis_pt, "Black", 22));
        docx = docx.add_paragraph(Self::build_paragraph(&anlysis_url_atch, "Black", 22));
        docx = docx.add_paragraph(Paragraph::new());
        docx = docx.add_paragraph(Self::build_paragraph(
            &"The Domain is clean as per virus total, Kaspersky and URL void.".to_string(),
            "Black",
            22,
        ));

        docx = docx.page_margin(PageMargin {
            top: 1440,    // 1 inch (in twentieths of a point)
            left: 1440,   // 1 inch
            bottom: 1440, // 1 inch
            right: 1440,  // 1 inch
            header: 1440, // 1 inch
            footer: 1440, // 1 inch
            // Gutter kept to 0 - it is increasing the margin on the left
            gutter: 0,
            // ..PageMargin::default()
        });

        return docx;
    }

    fn get_values(key: &str, map: &HashMap<String, String>) -> String {
        match map.get(key) {
            Some(v) => v.to_owned(),
            None => format!("NA"),
        }
    }

    fn head(text: &String, color: &str, size: usize) -> Paragraph {
        Paragraph::new()
            .add_run(
                Run::new()
                    .add_text(text)
                    .color(color)
                    .size(size)
                    .fonts(RunFonts::new().ascii("Open Sans")),
            )
            .align(docx_rs::AlignmentType::Center)
    }

    fn build_paragraph(text: &String, color: &str, size: usize) -> Paragraph {
        Paragraph::new().add_run(
            Run::new()
                .add_text(text)
                .color(color)
                .size(size)
                .fonts(RunFonts::new().ascii("Open Sans")),
        )
    }
}
