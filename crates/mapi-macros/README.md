# outlook-mapi-macros
This crate separates private `macro_rules!` macros used by the [outlook-mapi](https://crates.io/crates/outlook-mapi) crate from the public macros which it exports as part of its API. Exported macros can only invoke other exported macros, and there's no way to mark an exported macro as private.

## Getting Started
This crate is only intended for use in [outlook-mapi](https://crates.io/crates/outlook-mapi).
