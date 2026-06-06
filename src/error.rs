use thiserror::Error;

#[derive(Error, Debug)]
pub enum Error {
    #[error("Failed to load type library")]
    LoadTypeLibraryError(#[from] windows_core::Error),
    #[error("Type library not loaded")]
    TypeLibNotLoaded,
    #[error("IO Error")]
    IoError(#[from] std::io::Error),
    #[error("dotnet not found. Please install .NET SDK.")]
    DotnetNotFound,
    #[error("dotnet build failed with exit code {0}")]
    BuildFailed(i32),
}
