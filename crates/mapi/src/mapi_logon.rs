use outlook_mapi_sys::Microsoft::Office::Outlook::MAPI::Win32::*;
use std::{iter, ptr, sync::Arc};
use windows::Win32::Foundation::*;
use windows_core::*;

use crate::mapi_initialize::Initialize;

#[derive(Default)]
pub struct Flags {
    pub allow_others: bool,
    pub bg_session: bool,
    pub explicit_profile: bool,
    pub extended: bool,
    pub force_download: bool,
    pub logon_ui: bool,
    pub new_session: bool,
    pub no_mail: bool,
    pub nt_service: bool,
    pub service_ui_always: bool,
    pub timeout_short: bool,
    pub unicode: bool,
    pub use_default: bool,
}

impl Into<u32> for Flags {
    fn into(self) -> u32 {
        let allow_others = if self.allow_others {
            MAPI_ALLOW_OTHERS
        } else {
            0
        };
        let bg_session = if self.bg_session { MAPI_BG_SESSION } else { 0 };
        let explicit_profile = if self.explicit_profile {
            MAPI_EXPLICIT_PROFILE
        } else {
            0
        };
        let extended = if self.extended { MAPI_EXTENDED } else { 0 };
        let force_download = if self.force_download {
            MAPI_FORCE_DOWNLOAD
        } else {
            0
        };
        let logon_ui = if self.logon_ui { MAPI_LOGON_UI } else { 0 };
        let new_session = if self.new_session {
            MAPI_NEW_SESSION
        } else {
            0
        };
        let no_mail = if self.no_mail { MAPI_NO_MAIL } else { 0 };
        let nt_service = if self.nt_service { MAPI_NT_SERVICE } else { 0 };
        let service_ui_always = if self.service_ui_always {
            MAPI_SERVICE_UI_ALWAYS
        } else {
            0
        };
        let timeout_short = if self.timeout_short {
            MAPI_TIMEOUT_SHORT
        } else {
            0
        };
        let unicode = if self.unicode { MAPI_UNICODE } else { 0 };
        let use_default = if self.use_default {
            MAPI_USE_DEFAULT
        } else {
            0
        };

        allow_others
            | bg_session
            | explicit_profile
            | extended
            | force_download
            | logon_ui
            | new_session
            | no_mail
            | nt_service
            | service_ui_always
            | timeout_short
            | unicode
            | use_default
    }
}

pub struct Session {
    _initialized: Arc<Initialize>,
    pub session: IMAPISession,
}

impl Session {
    pub fn new(
        initialized: Arc<Initialize>,
        ui_param: HWND,
        profile_name: Option<&str>,
        password: Option<&str>,
        flags: Flags,
    ) -> Result<Self> {
        let mut profile_name: Option<Vec<_>> =
            profile_name.map(|value| value.bytes().chain(iter::once(0)).collect());
        let profile_name = profile_name
            .as_mut()
            .map(|value| value.as_mut_ptr())
            .unwrap_or(ptr::null_mut());
        let mut password: Option<Vec<_>> =
            password.map(|value| value.bytes().chain(iter::once(0)).collect());
        let password = password
            .as_mut()
            .map(|value| value.as_mut_ptr())
            .unwrap_or(ptr::null_mut());

        Ok(Self {
            _initialized: initialized,
            session: unsafe {
                let mut session = None;
                MAPILogonEx(
                    ui_param.0 as usize,
                    profile_name as *mut _,
                    password as *mut _,
                    flags.into(),
                    ptr::from_mut(&mut session),
                )?;
                session
            }
            .ok_or_else(|| Error::from(E_FAIL))?,
        })
    }
}
