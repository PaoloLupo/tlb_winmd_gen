use super::error;
use std::fs::File;
use std::io::Write;
use std::path::Path;
use std::process::Command;

pub fn generate_proj(path: &Path, lib_name: &str, winmd_dir: &Path) -> Result<(), error::Error> {
    let winmd_abs_path = std::fs::canonicalize(winmd_dir).unwrap_or(winmd_dir.to_path_buf());
    let winmd_file_path = winmd_abs_path.join(format!("{}.winmd", lib_name));

    let content = format!(
        r#"<?xml version="1.0" encoding="utf-8"?>
<Project Sdk="Microsoft.Windows.WinmdGenerator/0.70.11-preview">
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

pub fn generate_main_cpp(path: &Path, lib_name: &str) -> Result<(), error::Error> {
    let content = format!(r#"#include "{}.h""#, lib_name);
    let mut file = File::create(path)?;
    file.write_all(content.as_bytes())?;
    Ok(())
}

pub fn check_dotnet() -> Result<(), error::Error> {
    if Command::new("dotnet").arg("--version").output().is_err() {
        return Err(error::Error::DotnetNotFound);
    }
    Ok(())
}

pub fn run_dotnet_build(proj_dir: &Path) -> Result<(), error::Error> {
    let status = Command::new("dotnet")
        .arg("build")
        .arg("generate.proj")
        .current_dir(proj_dir)
        .status()?;

    if !status.success() {
        return Err(error::Error::BuildFailed(
            status.code().unwrap_or(-1),
        ));
    }
    Ok(())
}
