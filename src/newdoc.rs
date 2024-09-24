use docx_rs::{
    AlignmentType, Docx, LineSpacing, PageMargin, Paragraph, Run, RunFonts, Table, TableBorders,
    TableCell, TableRow,
};
use std::{
    collections::HashMap,
    fs::{self},
    path::{Path, PathBuf},
};

pub struct NewDocx {
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
        let from_address = Self::get_values("From", &headers);

        // Extracting sender domain
        let parts: Vec<&str> = from_address.trim().split("@").collect();
        let sender_domain: String = if parts.len() != 2 {
            "Not able to extract domain".to_string()
        } else {
            match parts.last() {
                Some(d) => d.replace(">", ""),
                None => "Not able to extract domain".to_string(),
            }
        };

        let mut have_attachments = false;
        let mut count = 0;
        for bh in &b_headers {
            let c_disposition = Self::get_values("Content-Disposition", &bh);
            if c_disposition != "NA" {
                have_attachments = true;
                count += 1;
            }
        }

        let attachments = if have_attachments {
            format!("Yes")
        } else {
            format!("No")
        };

        let date = Self::get_values("Date", &headers);
        let subject = Self::get_values("Subject", &headers);
        let to = Self::get_values("To", &headers);
        let return_path = Self::get_values(Self::RETURN_PATH, &headers);
        let h_content_type = Self::get_values(Self::CONTENT_TYPE, &headers);
        let spf = Self::get_values(Self::SPF, &headers);
        let auth_result = Self::get_values(Self::AUTH_RESULTS, &headers);
        let mut docx = Docx::new();

        let heading = &format!("{} {}", Self::HEADING, &self.i_number);
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::head(heading, Self::DARK_BLUE, Self::HEADING_SIZE))
                .align(AlignmentType::Center)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Self::build_paragraph(Self::SE_LINE, Self::DARK_BLUE, Self::REGULAR_SIZE)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_table(
            Table::new(vec![
                Self::table_row(Self::M_DATE, &date, Self::DEFAULT_BLACK, Self::REGULAR_SIZE),
                Self::table_row(
                    Self::SUBJECT,
                    &subject,
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ),
                Self::table_row(
                    Self::SENDER,
                    &from_address,
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ),
                Self::table_row(
                    Self::RECIPIENT,
                    &to,
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ),
                Self::table_row(
                    Self::DOMAIN,
                    &sender_domain,
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ),
                Self::table_row(
                    Self::BLK_LIST,
                    "No",
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ),
                Self::table_row(
                    Self::EML_GTWY,
                    "Delivered",
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ),
                Self::table_row(
                    Self::ATTACHMENTS,
                    &attachments,
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ),
                Self::table_row(Self::A_MAL, "****", Self::DEFAULT_BLACK, Self::REGULAR_SIZE),
                Self::table_row(Self::URL, "****", Self::DEFAULT_BLACK, Self::REGULAR_SIZE),
                Self::table_row(
                    Self::URL_MAL,
                    "****",
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ),
            ])
            .set_borders(TableBorders::new().clear_all()),
        );

        docx = docx.add_paragraph(Paragraph::new().line_spacing(LineSpacing::new().after(200)));

        docx = docx.add_paragraph(Self::build_paragraph(
            Self::REF,
            Self::DARK_BLUE,
            Self::SIDE_HEAD_SIZE,
        ));

        docx = docx.add_paragraph(
            Self::build_paragraph("*****", Self::RED, Self::REGULAR_SIZE)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Self::build_paragraph(
                Self::DOMAIN_ANALYSIS_HEAD,
                Self::DARK_BLUE,
                Self::SIDE_HEAD_SIZE,
            )
            .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Self::build_paragraph("*****", Self::RED, Self::REGULAR_SIZE)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Self::build_paragraph(Self::ANALYSIS_HEAD, Self::DARK_BLUE, Self::SIDE_HEAD_SIZE)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::build_run(
                    Self::ANALYSIS_VEC[0],
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .add_run(Self::build_run(
                    &from_address,
                    Self::RED,
                    Self::REGULAR_SIZE,
                ))
                .add_run(Self::build_run(
                    Self::ANALYSIS_VEC[1],
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .add_run(Self::build_run(
                    &sender_domain,
                    Self::RED,
                    Self::REGULAR_SIZE,
                ))
                .line_spacing(LineSpacing::new().after(200)),
        );

        let att_count = if count == 0 {
            String::from("no")
        } else {
            count.to_string()
        };

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::build_run(
                    Self::URL_ATTACHMENTS[0],
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .add_run(Self::build_run(
                    &att_count,
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .add_run(Self::build_run(
                    Self::URL_ATTACHMENTS[1],
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Self::build_paragraph(Self::DOMAIN_REP, Self::DEFAULT_BLACK, Self::REGULAR_SIZE)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::build_run(Self::REF, Self::RED, Self::REGULAR_SIZE))
                .add_run(Self::build_run(
                    Self::DOMAIN_REP_RES[0],
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .add_run(Self::build_run(
                    &sender_domain.to_ascii_lowercase(),
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::build_run(Self::REF, Self::RED, Self::REGULAR_SIZE))
                .add_run(Self::build_run(
                    Self::DOMAIN_REP_RES[1],
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .add_run(Self::build_run(
                    &sender_domain.to_ascii_lowercase(),
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::build_run(Self::REF, Self::RED, Self::REGULAR_SIZE))
                .add_run(Self::build_run(
                    Self::DOMAIN_REP_RES[2],
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .add_run(Self::build_run(
                    &sender_domain.to_ascii_lowercase(),
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Self::build_paragraph(Self::ANS_REPORT, Self::DARK_BLUE, Self::REGULAR_SIZE)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Self::build_paragraph(Self::VERDICT_HEAD, Self::DARK_BLUE, Self::SIDE_HEAD_SIZE)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::build_run(
                    Self::VERDICT_LINE[0],
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .add_run(Self::build_run(
                    Self::VERDICT_LINE[1],
                    Self::RED,
                    Self::REGULAR_SIZE,
                ))
                .add_run(Self::build_run(
                    Self::VERDICT_LINE[2],
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Self::build_paragraph(Self::SCREEN_SHOT, Self::DARK_BLUE, Self::SIDE_HEAD_SIZE)
                .line_spacing(LineSpacing::new().after(200)),
        );

        // docx = docx.add_paragraph();

        docx = docx.add_paragraph(Paragraph::new().line_spacing(LineSpacing::new().after(200)));
        docx = docx.add_paragraph(Paragraph::new().line_spacing(LineSpacing::new().after(200)));

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::head(
                    Self::HEADERS,
                    Self::DARK_BLUE,
                    Self::HEADING_SIZE,
                ))
                .align(AlignmentType::Center)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::build_run(
                    Self::AUTH_RESULTS,
                    Self::DARK_BLUE,
                    Self::SIDE_HEAD_SIZE,
                ))
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Self::build_paragraph(&auth_result, Self::DEFAULT_BLACK, Self::REGULAR_SIZE)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::build_run(
                    Self::RETURN_PATH,
                    Self::DARK_BLUE,
                    Self::SIDE_HEAD_SIZE,
                ))
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Self::build_paragraph(&return_path.trim(), Self::DEFAULT_BLACK, Self::REGULAR_SIZE)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::build_run(
                    Self::SPF,
                    Self::DARK_BLUE,
                    Self::SIDE_HEAD_SIZE,
                ))
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Self::build_paragraph(&spf, Self::DEFAULT_BLACK, Self::REGULAR_SIZE)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::build_run(
                    Self::CONTENT_TYPE,
                    Self::DARK_BLUE,
                    Self::SIDE_HEAD_SIZE,
                ))
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Self::build_paragraph(&h_content_type, Self::DEFAULT_BLACK, Self::REGULAR_SIZE)
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::build_run(
                    Self::B_CTYPE,
                    Self::DARK_BLUE,
                    Self::SIDE_HEAD_SIZE,
                ))
                .line_spacing(LineSpacing::new().after(200)),
        );

        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Self::build_run(
                    &format!("{:?}", &b_headers),
                    Self::DEFAULT_BLACK,
                    Self::REGULAR_SIZE,
                ))
                .line_spacing(LineSpacing::new().after(200)),
        );

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

    fn table_row(side_head: &str, main_data: &str, color: &str, size: usize) -> TableRow {
        TableRow::new(vec![
            TableCell::new().add_paragraph(
                Self::build_paragraph(side_head, color, size)
                    .line_spacing(LineSpacing::new().after(100)),
            ),
            TableCell::new().add_paragraph(
                Self::build_paragraph(&format!("{}{}", Self::MID_CELL, main_data), color, size)
                    .line_spacing(LineSpacing::new().after(100)),
            ),
        ])
    }

    fn get_values(key: &str, map: &HashMap<String, String>) -> String {
        match map.get(key) {
            Some(v) => v.to_owned(),
            None => format!("NA"),
        }
    }

    fn head(text: &str, color: &str, size: usize) -> Run {
        Run::new()
            .add_text(text)
            .color(color)
            .size(size)
            .fonts(RunFonts::new().ascii("Open Sans"))
        // .align(docx_rs::AlignmentType::Center)
    }

    fn build_paragraph(text: &str, color: &str, size: usize) -> Paragraph {
        Paragraph::new().add_run(Self::build_run(text, color, size))
    }

    fn build_run(text: &str, color: &str, size: usize) -> Run {
        Run::new()
            .add_text(text)
            .color(color)
            .size(size)
            .fonts(RunFonts::new().ascii("Open Sans"))
    }
}

impl NewDocx {
    const DARK_BLUE: &'static str = "#1D076D";
    const RED: &'static str = "#FF0000";
    const DEFAULT_BLACK: &'static str = "#000000";

    const HEADING_SIZE: usize = 36;
    const SIDE_HEAD_SIZE: usize = 28;
    const REGULAR_SIZE: usize = 22;

    const HEADING: &'static str = "INCIDENT";
    const SE_LINE: &'static str = "Please find the following initial analysis details.";
    const MID_CELL: &'static str = ":\t";

    const M_DATE: &'static str = "1. Date";
    const SUBJECT: &'static str = "2. Subject";
    const SENDER: &'static str = "3. Sender Id";
    const RECIPIENT: &'static str = "4. Recipient Id";
    const DOMAIN: &'static str = "5. Domain";
    const BLK_LIST: &'static str = "6. Blacklisted(Y/N)";
    const EML_GTWY: &'static str = "7. Email Gateway";
    const ATTACHMENTS: &'static str = "8. Attachments";
    const A_MAL: &'static str = "9. Attachments (Malicious)";
    const URL: &'static str = "10. URL(S)";
    const URL_MAL: &'static str = "11. URL (Malicious)";

    const REF: &'static str = "Ref: ";
    const DOMAIN_ANALYSIS_HEAD: &'static str = "Domain Analysis";
    const ANALYSIS_HEAD: &'static str = "Analysis";
    const ANALYSIS_VEC: [&'static str; 2] = [
        "User received a mail from ",
        " which was detected as a ***-suspicious mail. As per the initial analysis we gathered that the mail came from ",
    ];

    const URL_ATTACHMENTS: [&'static str; 2] = [
        "We also observed that there are *** URL(s) and ",
        " Attachment(s) in this email body.",
    ];

    const DOMAIN_REP: &'static str =
        "The Domain is clean as per virus total, Kaspersky and URL void.";
    const DOMAIN_REP_RES: [&'static str; 3] = [
        "https://www.urlvoid.com/scan/",
        "https://www.virustotal.com/gui/domain/",
        "https://talosintelligence.com/reputation_center/lookup?search=",
    ];
    const VERDICT_HEAD: &'static str = "Security Team verdict";
    const VERDICT_LINE: [&'static str; 3] = [
        "\tAs per our Analysis, we have reached a verdict that the attached email is ",
        "***** ",
        "Mail.",
    ];
    const SCREEN_SHOT: &'static str = "Screenshots:";

    const HEADERS: &'static str = "Mail - Headers";
    const RETURN_PATH: &'static str = "Return-Path";
    const CONTENT_TYPE: &'static str = "Content-Type";
    const SPF: &'static str = "Received-Spf";
    const AUTH_RESULTS: &'static str = "Authentication-Results";
    const B_CTYPE: &'static str = "Body Content-Type";

    const ANS_REPORT: &'static str = "As per the analysis we observed that, there is ******** Attached file is an html document and is trying to get the credentials of the user. Intention of the mail is credential harvesting.";
}
