use windows::Win32::System::Com::*;

mod types;
pub use types::*;

mod idispatch_ext;
mod utils;
mod variant_ext;

pub type Result<T> = ::windows::core::Result<T>;

pub struct SAPComInstance;

impl SAPComInstance {
    /// Initialise the COM environment.
    pub fn new() -> Result<Self> {
        log::debug!("CoInitialize'ing.");
        unsafe {
            CoInitialize(None)?;
        }
        Ok(SAPComInstance)
    }

    // Create an instance of the SAP wrapper
    pub fn sap_wrapper(&self) -> Result<SAPWrapper> {
        log::debug!("New CSapROTWrapper object generating.");
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
