use std::mem::ManuallyDrop;

use windows::{
    core::{self, BSTR},
    Win32::Foundation::VARIANT_BOOL,
    Win32::System::{
        Com::{IDispatch, VARIANT, VARIANT_0_0, VT_BOOL, VT_BSTR, VT_I4, VT_NULL},
        Ole::{VariantChangeType, VariantClear},
    },
};

use crate::{SAPComponent, types::GuiComponent};

pub(crate) trait VariantExt {
    fn null() -> VARIANT;
    fn from_i32(n: i32) -> VARIANT;
    fn from_i64(n: i64) -> VARIANT;
    fn from_str(s: &str) -> VARIANT;
    fn from_bool(b: bool) -> VARIANT;
    fn to_i32(&self) -> core::Result<i32>;
    fn to_i64(&self) -> core::Result<i64>;
    fn to_string(&self) -> core::Result<String>;
    fn to_bool(&self) -> core::Result<bool>;
    fn to_idispatch(&self) -> core::Result<&IDispatch>;
    fn to_sap_component(&self) -> core::Result<SAPComponent>;
}

impl VariantExt for VARIANT {
    fn null() -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0::default();
        v00.vt = VT_NULL;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_i32(n: i32) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0::default();
        v00.vt = VT_I4;
        v00.Anonymous.lVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_i64(n: i64) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0::default();
        v00.vt = VT_I4;
        v00.Anonymous.llVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_str(s: &str) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0::default();
        v00.vt = VT_BSTR;
        let bstr = BSTR::from(s);
        v00.Anonymous.bstrVal = ManuallyDrop::new(bstr);
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_bool(b: bool) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0::default();
        v00.vt = VT_BOOL;
        v00.Anonymous.boolVal = VARIANT_BOOL::from(b);
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn to_i32(&self) -> core::Result<i32> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, 0, VT_I4)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.lVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_i64(&self) -> core::Result<i64> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, 0, VT_I4)?;
            let v00 = &new.Anonymous.Anonymous;
            let n = v00.Anonymous.llVal;
            VariantClear(&mut new)?;
            Ok(n)
        }
    }
    fn to_string(&self) -> core::Result<String> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, 0, VT_BSTR)?;
            let v00 = &new.Anonymous.Anonymous;
            let str = v00.Anonymous.bstrVal.to_string();
            VariantClear(&mut new)?;
            Ok(str)
        }
    }
    fn to_bool(&self) -> core::Result<bool> {
        unsafe {
            let mut new = VARIANT::default();
            VariantChangeType(&mut new, self, 0, VT_BOOL)?;
            let v00 = &new.Anonymous.Anonymous;
            let b = v00.Anonymous.boolVal.as_bool();
            VariantClear(&mut new)?;
            Ok(b)
        }
    }
    fn to_idispatch(&self) -> core::Result<&IDispatch> {
        unsafe {
            let v00 = &self.Anonymous.Anonymous;
            let idisp = v00.Anonymous.pdispVal.as_ref().unwrap();
            Ok(idisp)
        }
    }
    fn to_sap_component(&self) -> core::Result<SAPComponent> {
        Ok(SAPComponent::from(GuiComponent {
            inner: self.to_idispatch()?.clone(),
        }))
    }
}
