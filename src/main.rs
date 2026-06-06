mod chm_doc;
#[cfg(feature = "standalone")]
mod embedded;
mod error;
mod flags;
mod idlgen;
mod project_gen;
mod ui;
use std::fs::{self, File};
use std::io::BufWriter;
use std::path::PathBuf;

use clap::Parser;

#[derive(Parser, Debug)]
#[command(version, about, long_about = None)]
struct Args {
    /// Path to the input TLB file
    tlb_path: PathBuf,

    /// Output directory for intermediate files (IDL, proj, cpp)
    #[arg(long, default_value = "proj")]
    out_dir: PathBuf,

    /// Output directory for the final .winmd file
    #[arg(long, default_value = "out")]
    winmd_dir: PathBuf,

    /// Run in TUI mode to inspect the TypeLib
    #[arg(long)]
    ui: bool,

    /// Path to CHM documentation file
    #[arg(long)]
    chm: Option<String>,

    /// Import stdole2.tlb in the generated IDL
    #[arg(long)]
    import_stdole: bool,
}

fn main() -> Result<(), error::Error> {
    #[cfg(feature = "standalone")]
    {
        let embedded = embedded::EmbeddedFiles::extract().map_err(|e| error::Error::IoError(e))?;
        if let Err(e) = ui::run(embedded.tlb_path, embedded.chm_path) {
            eprintln!("Error running TUI: {}", e);
            std::process::exit(1);
        }
        return Ok(());
    }

    #[cfg(not(feature = "standalone"))]
    {
        let args = Args::parse();
        let tlb_path = std::path::Path::new(&args.tlb_path);

        if args.ui {
            if let Err(e) = ui::run(tlb_path.to_path_buf(), args.chm) {
                eprintln!("Error running TUI: {}", e);
                std::process::exit(1);
            }
            return Ok(());
        }

        let out_dir = &args.out_dir;
        let winmd_dir = &args.winmd_dir;

        // Ensure output directories exist
        fs::create_dir_all(out_dir)?;
        fs::create_dir_all(winmd_dir)?;

        // Get library name from TLB
        let lib_name = idlgen::get_library_name(tlb_path)?;
        println!("Library Name: {}", lib_name);

        // Generate IDL
        let idl_path = out_dir.join(format!("{}.idl", lib_name));
        println!("Generating IDL: {}", idl_path.display());
        {
            let file = File::create(&idl_path)?;
            let mut writer = BufWriter::new(file);
            idlgen::build_tlb(tlb_path, &mut writer, args.import_stdole)?;
        }

        let proj_path = out_dir.join("generate.proj");
        println!("Generating Project File: {}", proj_path.display());
        project_gen::generate_proj(&proj_path, &lib_name, winmd_dir)?;

        let main_cpp_path = out_dir.join("main.cpp");
        println!("Generating main.cpp: {}", main_cpp_path.display());
        project_gen::generate_main_cpp(&main_cpp_path, &lib_name)?;

        // Check for dotnet
        if !project_gen::check_dotnet() {
            eprintln!("Error: 'dotnet' command not found. Please install .NET SDK.");
            return Ok(());
        }

        println!("Running dotnet build...");
        project_gen::run_dotnet_build(out_dir)?;

        println!("WinMD generation complete.");
        Ok(())
    }
}
