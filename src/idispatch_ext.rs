use windows::{
    core::{Result, GUID},
    Win32::System::Com::{
        IDispatch, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, VARIANT,
    },
};

use crate::{utils, variant_ext::VariantExt};

pub(crate) trait IDispatchExt {
    /// Call a function on this IDispatch
    fn call<S>(&self, name: S, args: Vec<VARIANT>) -> Result<VARIANT>
    where
        S: AsRef<str>;

    /// Get the value of a variable on this IDispatch
    fn get<S>(&self, name: S) -> Result<VARIANT>
    where
        S: AsRef<str>;

    /// Get the value of a variable on this IDispatch
    fn get_named<S>(&self, var_name: S, child_name: VARIANT) -> Result<VARIANT>
    where
        S: AsRef<str>;

    /// Set a value of a variable on this IDispatch
    fn set<S>(&self, name: S, value: VARIANT) -> Result<VARIANT>
    where
        S: AsRef<str>;

    /// Set a value of a variable on this IDispatch
    fn set_named<S>(&self, var_name: S, child_name: VARIANT, value: VARIANT) -> Result<VARIANT>
    where
        S: AsRef<str>;
}

impl IDispatchExt for IDispatch {
    fn call<S>(&self, name: S, args: Vec<VARIANT>) -> Result<VARIANT>
    where
        S: AsRef<str>,
    {
        let iid_null = GUID::zeroed();
        let mut result = VARIANT::null();
        unsafe {
            self.Invoke(
                utils::get_method_dispid(self, name)?,
                &iid_null,
                0,
                DISPATCH_METHOD,
                &utils::assemble_dispparams_get(args),
                Some(&mut result),
                None,
                None,
            )?;
        }
        Ok(result)
    }

    fn get<S>(&self, name: S) -> Result<VARIANT>
    where
        S: AsRef<str>,
    {
        let iid_null = GUID::zeroed();
        let mut result = VARIANT::null();
        unsafe {
            self.Invoke(
                utils::get_method_dispid(self, name)?,
                &iid_null,
                0,
                DISPATCH_PROPERTYGET,
                &DISPPARAMS::default(),
                Some(&mut result),
                None,
                None,
            )?;
        }
        Ok(result)
    }

    fn get_named<S>(&self, var_name: S, child_name: VARIANT) -> Result<VARIANT>
    where
        S: AsRef<str>,
    {
        let iid_null = GUID::zeroed();
        let mut result = VARIANT::null();
        unsafe {
            self.Invoke(
                utils::get_method_dispid(self, var_name)?,
                &iid_null,
                0,
                DISPATCH_PROPERTYGET,
                &utils::assemble_dispparams_get(vec![child_name]),
                Some(&mut result),
                None,
                None,
            )?;
        }
        Ok(result)
    }

    fn set<S>(&self, name: S, value: VARIANT) -> Result<VARIANT>
    where
        S: AsRef<str>,
    {
        let iid_null = GUID::zeroed();
        let mut result = VARIANT::null();
        unsafe {
            self.Invoke(
                utils::get_method_dispid(self, name)?,
                &iid_null,
                0,
                DISPATCH_PROPERTYPUT,
                &utils::assemble_dispparams_put(vec![value]),
                Some(&mut result),
                None,
                None,
            )?;
        }
        Ok(result)
    }

    fn set_named<S>(&self, var_name: S, child_name: VARIANT, value: VARIANT) -> Result<VARIANT>
    where
        S: AsRef<str>,
    {
        let iid_null = GUID::zeroed();
        let mut result = VARIANT::null();
        unsafe {
            self.Invoke(
                utils::get_method_dispid(self, var_name)?,
                &iid_null,
                0,
                DISPATCH_PROPERTYPUT,
                &utils::assemble_dispparams_put(vec![child_name, value]),
                Some(&mut result),
                None,
                None,
            )?;
        }
        Ok(result)
    }
}
