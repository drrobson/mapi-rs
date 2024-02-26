use windows::Win32::{Foundation::*, System::LibraryLoader::*};

fn get_mapi_module() -> HMODULE {
    use std::sync::OnceLock;
    use windows_core::*;

    static MAPI_MODULE: OnceLock<HMODULE> = OnceLock::new();
    *MAPI_MODULE.get_or_init(|| unsafe {
        #[cfg(feature = "olmapi32")]
        {
            GetModuleHandleW(w!("olmapi32")).expect("olmapi32 should already be loaded")
        }
        #[cfg(not(feature = "olmapi32"))]
        {
            LoadLibraryW(w!("mapi32")).expect("mapi32 should be loaded on demand")
        }
    })
}

#[macro_use]
extern crate outlook_mapi_stub;

#[allow(non_snake_case)]
pub mod Microsoft;
