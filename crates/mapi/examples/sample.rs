use outlook_mapi::*;
use std::sync::Arc;

fn main() {
    println!("Initializing MAPI...");
    let initialized =
        mapi_initialize::Initialize::new(Default::default()).expect("failed to initialize MAPI");
    println!("Trying to logon to the default profile...");
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
    println!("Success!");
}
