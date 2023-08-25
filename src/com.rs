use windows::{
    core::{self, ComInterface, GUID, HSTRING, PCWSTR},
    Win32::System::{
        Com::{
            CLSIDFromString, CoCreateInstance, IDispatch, CLSCTX_ALL, CLSCTX_LOCAL_SERVER,
            DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT,
            DISPPARAMS, VARIANT,
        },
        Ole::{GetActiveObject, DISPID_PROPERTYPUT},
    },
};

const LOCALE_USER_DEFAULT: u32 = 0x400;
const LOCALE_SYSTEM_DEFAULT: u32 = 0x0800;

pub struct ComObject {
    disp: IDispatch,
}

#[allow(unused)]
impl ComObject {
    /// COMオブジェクトを新規に作成します
    ///
    /// ProgIDかCLSID文字列 ( {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} 形式) を渡す
    pub fn new(id: &str) -> core::Result<Self> {
        unsafe {
            let lpsz = HSTRING::from(id);
            let rclsid = CLSIDFromString(&lpsz)?;
            let disp = match CoCreateInstance(&rclsid, None, CLSCTX_ALL) {
                Ok(disp) => disp,
                Err(_) => CoCreateInstance(&rclsid, None, CLSCTX_LOCAL_SERVER)?,
            };
            Ok(Self { disp })
        }
    }
    /// 起動中のExcelを捕まえるなどで使う
    ///
    /// ProgIDかCLSID文字列を渡す
    pub fn get(id: &str) -> core::Result<Option<Self>> {
        unsafe {
            let lpsz = HSTRING::from(id);
            let rclsid = CLSIDFromString(&lpsz)?;
            let pvreserved = std::ptr::null_mut();
            let mut ppunk = None;
            GetActiveObject(&rclsid, pvreserved, &mut ppunk)?;
            let disp: Option<IDispatch> = match ppunk {
                Some(unk) => {
                    // ComInterfaceをuseしておくことでcastが使える
                    let disp = unk.cast()?;
                    Some(disp)
                }
                None => None,
            };
            Ok(disp.map(|disp| Self { disp }))
        }
    }
    fn get_id_from_name(&self, name: &str) -> core::Result<i32> {
        unsafe {
            let hstring = HSTRING::from(name);
            let rgsznames = PCWSTR::from_raw(hstring.as_ptr());
            let mut rgdispid = 0;
            self.disp.GetIDsOfNames(
                &GUID::zeroed(),
                &rgsznames,
                1,
                LOCALE_USER_DEFAULT,
                &mut rgdispid,
            )?;
            Ok(rgdispid)
        }
    }
    fn invoke(
        &self,
        dispidmember: i32,
        pdispparams: &DISPPARAMS,
        wflags: DISPATCH_FLAGS,
    ) -> core::Result<VARIANT> {
        unsafe {
            let mut result = VARIANT::default();
            self.disp.Invoke(
                dispidmember,
                &GUID::zeroed(),
                LOCALE_SYSTEM_DEFAULT,
                wflags,
                pdispparams,
                Some(&mut result),
                None,
                None,
            )?;
            Ok(result)
        }
    }
    /// プロパティの値を得ます
    ///
    /// 値を得たいプロパティの名前を渡してください
    /// パラメータ付きプロパティの場合はパラメータを示すVARIANTを渡します
    pub fn get_property(&self, prop: &str, param: Option<VARIANT>) -> core::Result<VARIANT> {
        let dispidmember = self.get_id_from_name(prop)?;
        let mut pdispparams = DISPPARAMS::default();
        let mut args = if let Some(param) = param {
            vec![param]
        } else {
            vec![]
        };
        pdispparams.cArgs = args.len() as u32;
        pdispparams.rgvarg = args.as_mut_ptr();
        self.invoke(dispidmember, &pdispparams, DISPATCH_PROPERTYGET)
    }
    /// プロパティに値をセットします
    ///
    /// プロパティ名と必要ならばそのパラメータとセットする値を渡します
    pub fn set_property(
        &self,
        prop: &str,
        param: Option<VARIANT>,
        value: VARIANT,
    ) -> core::Result<()> {
        let dispidmember = self.get_id_from_name(prop)?;
        let mut pdispparams = DISPPARAMS::default();
        let mut args = if let Some(param) = param {
            vec![param, value]
        } else {
            vec![value]
        };
        let mut named_args = vec![DISPID_PROPERTYPUT];
        pdispparams.cArgs = args.len() as u32;
        pdispparams.rgvarg = args.as_mut_ptr();
        pdispparams.cNamedArgs = 1;
        pdispparams.rgdispidNamedArgs = named_args.as_mut_ptr();
        self.invoke(dispidmember, &pdispparams, DISPATCH_PROPERTYPUT)?;
        Ok(())
    }
    /// メソッドを実行します
    ///
    /// メソッド名とメソッドに渡す引数を渡します
    pub fn invoke_method(&self, method: &str, mut args: Vec<VARIANT>) -> core::Result<VARIANT> {
        let dispidmember = self.get_id_from_name(method)?;
        let mut pdispparams = DISPPARAMS::default();
        args.reverse();
        pdispparams.cArgs = args.len() as u32;
        pdispparams.rgvarg = args.as_mut_ptr();
        self.invoke(dispidmember, &pdispparams, DISPATCH_METHOD)
    }
}
