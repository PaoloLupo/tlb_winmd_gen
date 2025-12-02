# tlb_winmd_gen (WIP)

Command line tool to generate Windows Metadata files (.winmd) from type libraries (.tlb).

## Description

This project automates the conversion of legacy COM definitions (TLB) to the modern WinMD format, facilitating interoperability with languages like Rust (using windows-bindgen crate).

Based on [Generating metadata for the windows crate](https://www.withinrafael.com/2023/01/18/generating-metadata-for-the-windows-crate/).


## Requirements

- Rust
- .NET SDK (required for `dotnet build` command)

## Usage

Run the tool providing the path to the input TLB file:

```bash
cargo run -- <path_to_tlb> [options]
```

### Options

- `tlb_path`: Path to the input TLB file (required).
- `--out-dir`: Directory for intermediate files (default: "proj").
- `--winmd-dir`: Directory for the final .winmd file (default: "out").

