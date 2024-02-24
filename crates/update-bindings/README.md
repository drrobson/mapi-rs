# update-bindings
This crate is a utility which will regenerate the bindings in [mapi-sys](https://crates.io/crates/mapi-sys).

## Windows Metadata
The Windows crate requires a Windows Metadata (`winmd`) file describing the API. The one used in this crate was generated with the [webview2-win32md](https://github.com/wravery/webview2-win32md) project.
