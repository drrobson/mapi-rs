pub use outlook_mapi_sys::Microsoft;

use outlook_mapi_sys::Microsoft::Office::Outlook::MAPI::Win32::*;
use std::{mem, ptr};
use windows_core::*;

pub struct Session(Option<IMAPISession>);

impl Session {
    pub fn new(use_default: bool) -> Result<Self> {
        Ok(Self(unsafe {
            MAPIInitialize(ptr::from_mut(&mut MAPIINIT {
                ulVersion: MAPI_INIT_VERSION,
                ulFlags: 0,
            }) as *mut _)?;
            let mut session = None;
            MAPILogonEx(
                0,
                ptr::null_mut(),
                ptr::null_mut(),
                MAPI_EXTENDED
                    | MAPI_UNICODE
                    | MAPI_LOGON_UI
                    | if use_default { MAPI_USE_DEFAULT } else { 0 },
                ptr::from_mut(&mut session),
            )?;
            session
        }))
    }
}

impl Drop for Session {
    fn drop(&mut self) {
        mem::drop(self.0.take());
        unsafe {
            MAPIUninitialize();
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn login() {
        println!("Trying to logon...");
        let _session = Session::new(true).expect("should be able to init and logon to MAPI");
        println!("Created the session...");
    }
}
