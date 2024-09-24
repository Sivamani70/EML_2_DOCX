use mailparse::parse_mail;
use std::{
    collections::HashMap,
    fs::{self},
    path::PathBuf,
};

pub struct Mail {
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
