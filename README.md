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
- `--ui`: Launch the interactive Text User Interface (TUI) to inspect the TypeLib.

## TUI Mode

The tool includes a TUI for exploring the contents of a Type Library.

```bash
cargo run -- <path_to_tlb> --ui
```

### Features

- **Type Browser**: Navigate through all types (Interfaces, Enums, CoClasses) in the library.
- **Structured View**: View methods and enum values in a formatted table.
- **IDL Preview**: Toggle to view the raw IDL representation.
- **Search**:
    - **Type Search**: Filter the list of types.
    - **Member Search**: Filter methods or enum values within the selected type (`Ctrl+F`).
    - **Global Search**: Search for any function or enum value across the entire library (`Ctrl+P`).

### Shortcuts

- `Tab`: Toggle between Structured View and IDL Preview
- `Ctrl+F`: Toggle search focus between Types and Members
- `Ctrl+P`: Open Global Search Popup
- `Esc`: Close popup or exit
- `q`: Exit


