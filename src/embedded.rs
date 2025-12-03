use std::env;
use std::fs::File;
use std::io::Write;
use std::path::PathBuf;

pub struct EmbeddedFiles {
    pub tlb_path: PathBuf,
    pub chm_path: Option<String>,
    _temp_dir: tempfile::TempDir,
}

impl EmbeddedFiles {
    pub fn extract() -> Result<Self, std::io::Error> {
        let tlb_bytes = include_bytes!("../example/ETABSv1.tlb");
        let chm_bytes = include_bytes!("../example/ETABSv1.chm");

        let temp_dir = tempfile::Builder::new().prefix("etabs_tui_").tempdir()?;

        let tlb_path = temp_dir.path().join("ETABSv1.tlb");
        let mut tlb_file = File::create(&tlb_path)?;
        tlb_file.write_all(tlb_bytes)?;

        let chm_path = temp_dir.path().join("ETABSv1.chm");
        let mut chm_file = File::create(&chm_path)?;
        chm_file.write_all(chm_bytes)?;

        Ok(EmbeddedFiles {
            tlb_path,
            chm_path: Some(chm_path.to_string_lossy().to_string()),
            _temp_dir: temp_dir,
        })
    }
}
