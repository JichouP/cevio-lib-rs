#[derive(Debug, thiserror::Error)]
#[error(transparent)]
pub struct CeVIOError(pub anyhow::Error);

pub type Result<T> = std::result::Result<T, CeVIOError>;
