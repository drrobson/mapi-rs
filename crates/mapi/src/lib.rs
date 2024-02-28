//! All of the safe wrappers added by this crate, as well as any macros, are exported from the root
//! module of this crate.
//!
//! All of the nested, unsafe types from
//! [outlook-mapi-sys](https://crates.io/crates/outlook-mapi-sys) are re-exported as the `sys`
//! module in this crate.

/// Re-export all of the unsafe bindings from the
/// [outlook-mapi-sys](https://crates.io/crates/outlook-mapi-sys) crate.
pub mod sys {
    pub use outlook_mapi_sys::Microsoft::Office::Outlook::MAPI::Win32::*;
}

pub mod mapi_initialize;
pub mod mapi_logon;
pub mod prop_value;
pub mod row;
pub mod row_set;

pub use mapi_initialize::*;
pub use mapi_logon::*;
pub use prop_value::*;
pub use row::*;
pub use row_set::*;

mod macros;
