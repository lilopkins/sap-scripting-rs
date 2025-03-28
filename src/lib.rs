//! SAP Scripting for Rust
//!
//! See the examples for how to use this library.

use windows::Win32::System::Com::*;

pub use com_shim::IsA;

/// The types from this library.
pub mod types;

pub use types::*;

/// A result of a call.
pub type Result<T> = ::windows::core::Result<T>;

/// An instance of a COM session. This should be kept whilst a connection to SAP is used.
pub struct SAPComInstance;

impl SAPComInstance {
    /// Initialise the COM environment.
    pub fn new() -> Result<Self> {
        tracing::debug!("CoInitialize'ing.");
        unsafe {
            CoInitialize(None)?;
        }
        Ok(SAPComInstance)
    }

    /// Create an instance of the SAP wrapper
    pub fn sap_wrapper(&self) -> Result<SAPWrapper> {
        tracing::debug!("New CSapROTWrapper object generating.");
        SAPWrapper::new()
    }
}

impl Drop for SAPComInstance {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}
