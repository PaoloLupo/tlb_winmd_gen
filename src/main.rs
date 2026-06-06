use std::fs::{self, File};
use std::io::BufWriter;
use std::path::PathBuf;

use tlb_winmd_gen::{error, idlgen, project_gen, ui};

use clap::{Parser, Subcommand};

#[derive(Parser, Debug)]
#[command(version, about = "TLB to WinMD generator - converts COM Type Libraries to Windows Metadata", long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand, Debug)]
enum Commands {
    /// Generate a .winmd file from a TLB file
    Generate(GenerateArgs),
    /// Inspect a TLB file using an interactive TUI
    Inspect(InspectArgs),
}

#[derive(Parser, Debug)]
struct GenerateArgs {
    /// Path to the input TLB file
    tlb_path: PathBuf,

    /// Output directory for intermediate files (IDL, proj, cpp)
    #[arg(long, default_value = "proj")]
    out_dir: PathBuf,

    /// Output directory for the final .winmd file
    #[arg(long, default_value = "out")]
    winmd_dir: PathBuf,

    /// Import stdole2.tlb in the generated IDL
    #[arg(long)]
    import_stdole: bool,
}

#[derive(Parser, Debug)]
struct InspectArgs {
    /// Path to the input TLB file
    tlb_path: PathBuf,

    /// Path to CHM documentation file
    #[arg(long)]
    chm: Option<String>,
}

fn main() -> Result<(), error::Error> {
    let cli = Cli::parse();

    match cli.command {
        Commands::Inspect(args) => {
            let tlb_path = std::path::Path::new(&args.tlb_path);
            if let Err(e) = ui::run(tlb_path.to_path_buf(), args.chm) {
                eprintln!("Error running TUI: {}", e);
                std::process::exit(1);
            }
            Ok(())
        }
        Commands::Generate(args) => {
            let tlb_path = std::path::Path::new(&args.tlb_path);
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
            project_gen::check_dotnet()?;

            println!("Running dotnet build...");
            project_gen::run_dotnet_build(out_dir)?;

            println!("WinMD generation complete.");
            Ok(())
        }
    }
}
