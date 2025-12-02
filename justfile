set windows-shell := ["powershell.exe", "-NoLogo", "-Command"]

run:
    cargo run --release -- "example/ETABSv1.tlb" --chm "./example/ETABSv1.chm" --ui

build:
    cargo build --release
