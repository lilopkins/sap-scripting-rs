use windows::{
    core::*,
    Win32::System::{Com::*, Ole::*},
};

pub(crate) fn get_method_dispid<S>(disp: &IDispatch, name: S) -> Result<i32>
where
    S: AsRef<str>,
{
    unsafe {
        let riid = GUID::zeroed();
        let hstring = HSTRING::from(name.as_ref());
        let rgsznames = PCWSTR::from_raw(hstring.as_ptr());
        let cnames = 1;
        let lcid = 0;
        let mut dispidmember = 0;

        disp.GetIDsOfNames(&riid, &rgsznames, cnames, lcid, &mut dispidmember)?;
        Ok(dispidmember)
    }
}

pub(crate) fn assemble_dispparams_get(mut args: Vec<VARIANT>) -> DISPPARAMS {
    DISPPARAMS {
        rgvarg: args.as_mut_ptr(),
        cArgs: args.len() as u32,
        ..Default::default()
    }
}

pub(crate) fn assemble_dispparams_put(mut args: Vec<VARIANT>) -> DISPPARAMS {
    let mut named_args = vec![DISPID_PROPERTYPUT];
    DISPPARAMS {
        rgvarg: args.as_mut_ptr(),
        cArgs: args.len() as u32,
        cNamedArgs: named_args.len() as u32,
        rgdispidNamedArgs: named_args.as_mut_ptr(),
    }
}
