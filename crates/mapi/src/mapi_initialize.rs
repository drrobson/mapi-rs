use crate::sys::*;
use core::ptr;
use windows_core::*;

#[derive(Default)]
pub struct InitializeFlags {
    pub multithread_notifications: bool,
    pub nt_service: bool,
    pub no_coinit: bool,
}

impl From<InitializeFlags> for u32 {
    fn from(value: InitializeFlags) -> Self {
        let multithread_notifications = if value.multithread_notifications {
            MAPI_MULTITHREAD_NOTIFICATIONS
        } else {
            0
        };
        let nt_service = if value.nt_service { MAPI_NT_SERVICE } else { 0 };
        let no_coinit = if value.no_coinit { MAPI_NO_COINIT } else { 0 };

        multithread_notifications | nt_service | no_coinit
    }
}

pub struct Initialize();

impl Initialize {
    pub fn new(flags: InitializeFlags) -> Result<Self> {
        unsafe {
            MAPIInitialize(ptr::from_mut(&mut MAPIINIT {
                ulVersion: MAPI_INIT_VERSION,
                ulFlags: flags.into(),
            }) as *mut _)?;
        }

        Ok(Self())
    }
}

impl Drop for Initialize {
    fn drop(&mut self) {
        unsafe {
            MAPIUninitialize();
        }
    }
}
