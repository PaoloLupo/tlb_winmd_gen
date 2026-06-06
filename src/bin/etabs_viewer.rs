use std::path::PathBuf;

fn main() {
    let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let tlb_path = manifest_dir.join("etabs").join("ETABSv1.tlb");
    let chm_path = manifest_dir.join("etabs").join("ETABSv1.chm");

    let chm_str = chm_path.to_str().map(|s| s.to_string());

    if let Err(e) = tlb_winmd_gen::ui::run(tlb_path, chm_str) {
        eprintln!("Error: {e}");
        std::process::exit(1);
    }
}
