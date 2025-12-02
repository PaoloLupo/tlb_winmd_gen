mod error;
mod idlgen;
use std::fs::{self, File};
use std::io::{BufWriter, Write};
use std::path::{Path, PathBuf};
use std::process::Command;

use clap::Parser;

#[derive(Parser, Debug)]
#[command(version, about, long_about = None)]
struct Args {
    /// Path to the input TLB file
    tlb_path: PathBuf,

    /// Output directory for intermediate files (IDL, proj, cpp)
    #[arg(long, default_value = "project")]
    out_dir: PathBuf,

    /// Output directory for the final .winmd file
    #[arg(long, default_value = "out")]
    winmd_dir: PathBuf,
}

fn main() -> Result<(), error::Error> {
    let args = Args::parse();
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
        idlgen::build_tlb(tlb_path, &mut writer)?;
    }

    // Generate generate.proj
    let proj_path = out_dir.join("generate.proj");
    println!("Generating Project File: {}", proj_path.display());
    generate_proj(&proj_path, &lib_name, winmd_dir)?;

    // Generate main.cpp
    let main_cpp_path = out_dir.join("main.cpp");
    println!("Generating main.cpp: {}", main_cpp_path.display());
    generate_main_cpp(&main_cpp_path, &lib_name)?;

    // Check for dotnet
    if !check_dotnet() {
        eprintln!("Error: 'dotnet' command not found. Please install .NET SDK.");
        return Ok(());
    }

    // Run dotnet build
    println!("Running dotnet build...");
    run_dotnet_build(out_dir)?;

    println!("WinMD generation complete.");
    Ok(())
}

fn generate_proj(path: &Path, lib_name: &str, winmd_dir: &Path) -> Result<(), error::Error> {
    // We need relative path from .metadata to .windows/winmd for the project file if possible,
    // or absolute path. Let's use absolute path for simplicity or relative if we can compute it.
    // For now, let's try to use the path provided in args, assuming it's relative to CWD.
    // The project file is in out_dir.

    // To make it robust, let's just use the path string provided by user/default,
    // but we need to adjust it because the project file is running from out_dir.
    // Actually, let's just use the absolute path of winmd_dir to avoid relative path hell.
    let winmd_abs_path = std::fs::canonicalize(winmd_dir).unwrap_or(winmd_dir.to_path_buf());
    let winmd_file_path = winmd_abs_path.join(format!("{}.winmd", lib_name));

    let content = format!(
        r#"<?xml version="1.0" encoding="utf-8"?>
<Project Sdk="Microsoft.Windows.WinmdGenerator/0.65.8-preview">
  <PropertyGroup Label="Globals">
    <OutputWinmd>{}</OutputWinmd>
    <WinmdVersion>255.255.255.255</WinmdVersion>
    <IdlsRoot>$(MSBuildThisFileDirectory)</IdlsRoot>
    <AdditionalIncludes>$(CompiledHeadersDir)</AdditionalIncludes>
  </PropertyGroup>
  <ItemGroup>
    <Idls Include="$(IdlsRoot)\{}.idl"/>
    <Headers Include="$(CompiledHeadersDir)\{}.h"/>
    <Partition Include="main.cpp">
      <TraverseFiles>@(Headers)</TraverseFiles>
      <Namespace>{}</Namespace>
    </Partition>
  </ItemGroup>
</Project>"#,
        winmd_file_path.display(),
        lib_name,
        lib_name,
        lib_name
    );

    let mut file = File::create(path)?;
    file.write_all(content.as_bytes())?;
    Ok(())
}

fn generate_main_cpp(path: &Path, lib_name: &str) -> Result<(), error::Error> {
    let content = format!(r#"#include "{}.h""#, lib_name);
    let mut file = File::create(path)?;
    file.write_all(content.as_bytes())?;
    Ok(())
}

fn check_dotnet() -> bool {
    Command::new("dotnet").arg("--version").output().is_ok()
}

fn run_dotnet_build(proj_dir: &Path) -> Result<(), error::Error> {
    let status = Command::new("dotnet")
        .arg("build")
        .arg("generate.proj")
        .current_dir(proj_dir)
        .status()?;

    if !status.success() {
        // Return a generic IO error for build failure, as we don't have a specific error variant for it yet
        return Err(error::Error::IoError(std::io::Error::new(
            std::io::ErrorKind::Other,
            "dotnet build failed",
        )));
    }
    Ok(())
}
