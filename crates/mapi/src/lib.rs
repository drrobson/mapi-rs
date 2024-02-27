pub use outlook_mapi_sys::Microsoft;

pub mod mapi_initialize;
pub mod mapi_logon;

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Arc;

    #[test]
    fn login() {
        let initialized = mapi_initialize::Initialize::new(Default::default())
            .expect("failed to initialize MAPI");
        let _session = mapi_logon::Session::new(
            Arc::new(initialized),
            Default::default(),
            None,
            None,
            mapi_logon::Flags {
                extended: true,
                unicode: true,
                logon_ui: true,
                use_default: true,
                ..Default::default()
            },
        )
        .expect("should be able to logon to the default MAPI profile");
    }
}
