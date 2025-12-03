use chmlib::ChmFile;
use chmlib::Filter;
use regex::Regex;
use std::collections::HashMap;
use std::path::Path;
use std::sync::{Arc, Mutex};

#[derive(Debug, Clone)]
pub struct ParamDoc {
    pub name: String,
    pub description: String,
}

#[derive(Debug, Clone)]
pub struct DocStructure {
    pub description: String,
    pub parameters: Vec<ParamDoc>,
}

pub struct ChmDocumentationProvider {
    chm_file: Arc<Mutex<ChmFile>>,
    index: HashMap<String, String>,
}

impl ChmDocumentationProvider {
    pub fn new(path: &str) -> Result<Self, chmlib::OpenError> {
        let mut chm = ChmFile::open(path)?;
        let mut index = HashMap::new();

        // Try to find and parse .hhk file (Index)
        let mut hhk_path = None;
        let _ = chm.for_each(Filter::all(), |_, unit| {
            if let Some(path) = unit.path() {
                if let Some(path_str) = path.to_str() {
                    if path_str.to_lowercase().ends_with(".hhk") {
                        hhk_path = Some(path_str.to_string());
                        return chmlib::Continuation::Stop;
                    }
                }
            }
            chmlib::Continuation::Continue
        });

        if let Some(p) = hhk_path {
            if let Some(unit) = chm.find(&p) {
                let mut buffer = vec![0; unit.length() as usize];
                if chm.read(&unit, 0, &mut buffer).is_ok() {
                    let content = String::from_utf8_lossy(&buffer);
                    index = parse_hhk(&content);
                }
            }
        }

        Ok(Self {
            chm_file: Arc::new(Mutex::new(chm)),
            index,
        })
    }

    pub fn get_doc(&self, name: &str) -> Option<DocStructure> {
        let mut chm = self.chm_file.lock().unwrap();
        let mut found_path = None;
        let search_name = name.to_lowercase();

        // 1. Try Index Lookup
        if !self.index.is_empty() {
            // Exact match
            for (key, path) in &self.index {
                if key.to_lowercase() == search_name {
                    found_path = Some(path.clone());
                    break;
                }
            }

            // Fuzzy match if exact match fails
            if found_path.is_none() {
                for (key, path) in &self.index {
                    let key_lower = key.to_lowercase();
                    if key_lower.contains(&search_name) {
                        found_path = Some(path.clone());
                        break;
                    }
                }
            }
        }

        // 2. Fallback to heuristic search
        if found_path.is_none() {
            let _ = chm.for_each(Filter::all(), |_, unit| {
                if let Some(path) = unit.path() {
                    if let Some(path_str) = path.to_str() {
                        let path_lower = path_str.to_lowercase();
                        let filename = Path::new(path_str)
                            .file_name()
                            .and_then(|n| n.to_str())
                            .unwrap_or("")
                            .to_lowercase();

                        if filename.contains(&search_name)
                            && (path_lower.ends_with(".htm") || path_lower.ends_with(".html"))
                        {
                            found_path = Some(path_str.to_string());
                            return chmlib::Continuation::Stop;
                        }
                    }
                }
                chmlib::Continuation::Continue
            });
        }

        if let Some(path) = found_path {
            let try_paths = vec![path.clone(), format!("/{}", path)];

            for p in try_paths {
                if let Some(unit) = chm.find(&p) {
                    let mut buffer = vec![0; unit.length() as usize];
                    if chm.read(&unit, 0, &mut buffer).is_ok() {
                        let html_content = String::from_utf8_lossy(&buffer);
                        return Some(parse_html(&html_content));
                    }
                }
            }
        }

        None
    }
}

fn parse_hhk(content: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let mut current_name = None;

    for line in content.lines() {
        let line = line.trim();
        if line.contains("name=\"Name\"") {
            if let Some(start) = line.find("value=\"") {
                let rest = &line[start + 7..];
                if let Some(end) = rest.find("\"") {
                    current_name = Some(rest[0..end].to_string());
                }
            }
        } else if line.contains("name=\"Local\"") {
            if let Some(name) = current_name.take() {
                if let Some(start) = line.find("value=\"") {
                    let rest = &line[start + 7..];
                    if let Some(end) = rest.find("\"") {
                        let local = rest[0..end].to_string();
                        map.insert(name, local);
                    }
                }
            }
        }
    }

    map
}

fn parse_html(html: &str) -> DocStructure {
    let mut description = String::new();
    let mut parameters = Vec::new();

    // Pre-process: Replace specific script for separator with "."
    // e.g. <script type="text/javascript">AddLanguageSpecificTextSet("...|nu=.");</script>
    let re_sep_script =
        Regex::new(r#"(?is)<script[^>]*>AddLanguageSpecificTextSet.*?</script>"#).unwrap();
    let html_processed = re_sep_script.replace_all(html, ".");

    // 2. Extract Description
    // Try <div class="summary"> first
    let re_summary = Regex::new(r"(?is)<div class=.summary.>(.*?)</div>").unwrap();
    if let Some(caps) = re_summary.captures(&html_processed) {
        description = strip_html_tags(&caps[1]);
    } else {
        // Fallback: text after title
        // Stop at <h, <div id="syntax", <div id="parameters", OR <div class="collapsibleAreaRegion">
        let re_desc = Regex::new(r"(?is)</h1>\s*(.*?)\s*(?:<h|<div id=.syntax.|<div id=.parameters.|<div class=.collapsibleAreaRegion.)").unwrap();
        if let Some(caps) = re_desc.captures(&html_processed) {
            let raw_desc = strip_html_tags(&caps[1]);
            // Filter out Namespace/Assembly lines which are common in this fallback area
            description = raw_desc
                .lines()
                .map(|line| line.trim())
                .filter(|line| {
                    !line.is_empty()
                        && !line.starts_with("Namespace:")
                        && !line.starts_with("Assembly:")
                        && !line.starts_with("Version:")
                        && !line.contains("(in ")
                        && !line.contains("ETABSv1") // Filter out library name artifacts
                })
                .collect::<Vec<&str>>()
                .join("\n");
        }
    }

    // 3. Extract Parameters (Structured)
    // Look for the parameters section first
    let re_params_section = Regex::new(r"(?is)(?:<h3[^>]*>|<h4[^>]*>|<strong>)Parameters(?:</h3>|</h4>|</strong>)(.*?)(?:<h3[^>]*>|<h4[^>]*>|<strong>|<div id=.remarks.|<div id=.example.|<div class=.collapsibleAreaRegion.)").unwrap();

    if let Some(section_caps) = re_params_section.captures(&html_processed) {
        let section_html = &section_caps[1];

        // Parse <dt>...<dd> pairs
        // <dt><span class="parameter">Name</span></dt>
        // <dd>Type: ... <br /> Description </dd>

        let re_dt = Regex::new(r"(?is)<dt>(.*?)</dt>").unwrap();
        let re_dd = Regex::new(r"(?is)<dd>(.*?)</dd>").unwrap();

        let dts: Vec<_> = re_dt.find_iter(section_html).collect();
        let dds: Vec<_> = re_dd.find_iter(section_html).collect();

        for (dt, dd) in dts.iter().zip(dds.iter()) {
            let dt_text = strip_html_tags(dt.as_str()); // Should be just the name
            let dd_html = dd.as_str();

            // Extract Type (usually starts with "Type: " and ends at <br />)
            let mut param_desc = String::new();

            let re_type = Regex::new(r"(?is)Type:\s*(.*?)(?:<br[^>]*>|$)").unwrap();
            if let Some(type_caps) = re_type.captures(dd_html) {
                // Description is everything after the type line
                // Find where the type match ended
                if let Some(m) = type_caps.get(0) {
                    let rest = &dd_html[m.end()..];
                    let raw_desc = strip_html_tags(rest);
                    // Clean up leading slash if present (artifact of malformed html or split tags)
                    param_desc = raw_desc.trim_start_matches('/').trim().to_string();
                }
            } else {
                // Fallback if "Type:" not found
                param_desc = strip_html_tags(dd_html);
            }

            parameters.push(ParamDoc {
                name: dt_text,
                description: param_desc,
            });
        }
    }

    DocStructure {
        description,
        parameters,
    }
}

fn strip_html_tags(html: &str) -> String {
    // First remove script and style blocks content
    let re_script = Regex::new(r"(?is)<script[^>]*>.*?</script>").unwrap();
    let no_script = re_script.replace_all(html, "");

    let re_style = Regex::new(r"(?is)<style[^>]*>.*?</style>").unwrap();
    let no_style = re_style.replace_all(&no_script, "");

    let mut result = String::new();
    let mut inside_tag = false;

    for c in no_style.chars() {
        if c == '<' {
            inside_tag = true;
        } else if c == '>' {
            inside_tag = false;
            result.push(' '); // Add space to separate text
        } else if !inside_tag {
            result.push(c);
        }
    }

    // Basic cleanup of multiple spaces and newlines
    result
        .lines()
        .map(|line| line.trim())
        .filter(|line| !line.is_empty())
        .collect::<Vec<&str>>()
        .join("\n")
}
