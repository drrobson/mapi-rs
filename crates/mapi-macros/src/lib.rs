//! Private macros used by the [outlook-mapi](https://crates.io/crates/outlook-mapi) crate
//! internally that are not re-exported.

/// Build the common casting function `impl` block for all of the SizedXXX macros.
#[macro_export]
macro_rules! impl_sized_struct_casts {
    ($name:ident, $sys_type:path) => {
        #[allow(dead_code)]
        impl $name {
            pub fn as_ptr(&self) -> *const $sys_type {
                unsafe { std::mem::transmute::<&Self, &$sys_type>(self) }
            }

            pub fn as_mut_ptr(&mut self) -> *mut $sys_type {
                unsafe { std::mem::transmute::<&mut Self, &mut $sys_type>(self) }
            }
        }
    };
}

/// Build an optional `impl Default` block for any of the SizedXXX macros.
#[macro_export]
macro_rules! impl_sized_struct_default {
    ($name:ident $body:tt) => {
        #[allow(dead_code)]
        impl Default for $name {
            fn default() -> Self {
                Self $body
            }
        }
    };
}
