use windows::{
    core::{self, BSTR},
    Win32::{
        Foundation::VARIANT_BOOL,
        System::{
            Com::{
                SAFEARRAY, VARENUM, VARIANT, VARIANT_0_0, VT_ARRAY, VT_BOOL, VT_BSTR, VT_BYREF,
                VT_I4, VT_NULL, VT_VARIANT,
            },
            Ole::{VariantChangeType, VariantClear},
        },
    },
};

use std::mem::ManuallyDrop;

pub trait VariantExt {
    /// VT_NULLなVARIANTを作る
    fn null() -> VARIANT;
    /// VT_BYREF|VT_VARIANTなVARIANTを作る、参照渡し用
    /// 引数は参照先となるVARIANT
    fn by_ref(var_val: *mut VARIANT) -> VARIANT;
    /// VT_I4を作る
    fn from_i32(n: i32) -> VARIANT;
    /// VT_BSTRを作る
    fn from_str(s: &str) -> VARIANT;
    /// VT_BOOLを作る
    fn from_bool(b: bool) -> VARIANT;
    /// VT_ARRAY|VT_VARIANTを作る
    fn from_safearray(psa: *mut SAFEARRAY) -> VARIANT;
    /// VARIANTをi32にする
    fn to_i32(&self) -> core::Result<i32>;
    /// VARIANTをStringにする
    fn to_string(&self) -> core::Result<String>;
    /// VARIANTをboolにする
    fn to_bool(&self) -> core::Result<bool>;
}

impl VariantExt for VARIANT {
    fn null() -> VARIANT {
        let mut variant = VARIANT::default();
        let v00 = VARIANT_0_0 {
            vt: VT_NULL,
            ..Default::default()
        };
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn by_ref(var_val: *mut VARIANT) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VARENUM(VT_BYREF.0 | VT_VARIANT.0),
            ..Default::default()
        };
        v00.Anonymous.pvarVal = var_val;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_i32(n: i32) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_I4,
            ..Default::default()
        };
        v00.Anonymous.lVal = n;
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_str(s: &str) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_BSTR,
            ..Default::default()
        };
        let bstr = BSTR::from(s);
        v00.Anonymous.bstrVal = ManuallyDrop::new(bstr);
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_bool(b: bool) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VT_BOOL,
            ..Default::default()
        };
        v00.Anonymous.boolVal = VARIANT_BOOL::from(b);
        variant.Anonymous.Anonymous = ManuallyDrop::new(v00);
        variant
    }
    fn from_safearray(psa: *mut SAFEARRAY) -> VARIANT {
        let mut variant = VARIANT::default();
        let mut v00 = VARIANT_0_0 {
            vt: VARENUM(VT_ARRAY.0 | VT_VARIANT.0),
            ..Default::default()
        };
        v00.Anonymous.parray = psa;
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
}
