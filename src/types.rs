use windows::{core::*, Win32::System::Com::*};

use crate::idispatch_ext::IDispatchExt;
use crate::variant_ext::VariantExt;

/// A wrapper over the SAP scripting engine, equivalent to CSapROTWrapper.
pub struct SAPWrapper {
    inner: IDispatch,
}

impl SAPWrapper {
    pub(crate) fn new() -> crate::Result<Self> {
        unsafe {
            let clsid: GUID = CLSIDFromProgID(w!("SapROTWr.SapROTWrapper"))?;
            let p_clsid: *const GUID = &clsid;
            log::debug!("CSapROTWrapper CLSID: {:?}", clsid);

            let dispatch: IDispatch =
                CoCreateInstance(p_clsid, None, CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER)?;
            Ok(SAPWrapper { inner: dispatch })
        }
    }

    /// Get the Scripting Engine object from this wrapper.
    pub fn scripting_engine(&self) -> crate::Result<GUIApplication> {
        log::debug!("Getting UI ROT entry...");
        let result = self
            .inner
            .call("GetROTEntry", vec![VARIANT::from_str("SAPGUI")])?;

        let sap_gui = result.to_idispatch()?;

        log::debug!("Getting scripting engine.");
        let scripting_engine = sap_gui.call("GetScriptingEngine", vec![])?;

        Ok(GUIApplication {
            inner: scripting_engine.to_idispatch()?.clone(),
        })
    }
}

pub struct GUIApplication {
    inner: IDispatch,
}

impl GUIApplication {
    /// Start a new connection with the provided connection string.
    pub fn start_connection<S>(&self, connection_string: S) -> crate::Result<GuiConnection>
    where
        S: AsRef<str>,
    {
        Ok(GuiConnection {
            inner: self
                .inner
                .call(
                    "OpenConnectionByConnectionString",
                    vec![VARIANT::from_str(connection_string.as_ref())],
                )?
                .to_idispatch()?
                .clone(),
        })
    }

    /// Get the nth session from the scripting engine.
    pub fn get_connection(&self, n: i32) -> crate::Result<GuiConnection> {
        Ok(GuiConnection {
            inner: self
                .inner
                .get_named("Children", VARIANT::from_i32(n))?
                .to_idispatch()?
                .clone(),
        })
    }
}

pub struct GuiConnection {
    inner: IDispatch,
}

impl GuiConnection {
    pub fn get_session(&self) -> crate::Result<GuiSession> {
        Ok(GuiSession {
            inner: self
                .inner
                .get_named("Children", VARIANT::from_i32(0))?
                .to_idispatch()?
                .clone(),
        })
    }
}

pub struct GuiSession {
    inner: IDispatch,
}

impl GuiSession {
    pub fn find_by_id<S>(&self, id: S) -> crate::Result<GuiComponent>
    where
        S: AsRef<str>,
    {
        Ok(GuiComponent {
            inner: self
                .inner
                .call("FindById", vec![VARIANT::from_str(id.as_ref())])?
                .to_idispatch()?
                .clone(),
        })
    }
}

pub struct GuiComponent {
    inner: IDispatch,
}

impl GuiComponent {
    pub fn kind(&self) -> crate::Result<String> {
        self.inner.get("Type")?.to_string()
    }
}
