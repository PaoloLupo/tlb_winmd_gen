use thiserror::Error;

#[derive(Error, Debug)]
pub enum Error {
    #[error("Failed to load type library")]
    LoadTypeLibraryError(#[from] windows_core::Error),
    #[error("Type library not loaded")]
    TypeLibNotLoaded,
    #[error("IO Error")]
    IoError(#[from] std::io::Error),
}
