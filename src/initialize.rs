pub struct Initialize {}

impl Initialize {
    pub fn new() -> anyhow::Result<Self> {
        use windows::Win32::System::Com::{
            CoInitializeEx, COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE,
        };
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE) }?;

        Ok(Self {})
    }
}

impl Drop for Initialize {
    fn drop(&mut self) {
        use windows::Win32::System::Com::CoUninitialize;
        unsafe { CoUninitialize() };
    }
}
