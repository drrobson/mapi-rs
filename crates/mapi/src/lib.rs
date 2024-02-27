pub use outlook_mapi_sys::Microsoft::Office::Outlook::MAPI::Win32 as sys;

pub mod mapi_initialize;
pub mod mapi_logon;
pub mod row;
pub mod row_set;

pub use mapi_initialize::*;
pub use mapi_logon::*;
pub use row::*;
pub use row_set::*;

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSPropTagArray {
    ($count:expr, $($tags:expr),+) => {
        {
            #[repr(C)]
            struct PropTagArray {
                count: u32,
                tags: [u32; $count],
            }
            std::mem::transmute::<_, &mut $crate::sys::SPropTagArray>(&mut PropTagArray {
                count: $count,
                tags: [$($tags,)+],
            })
        }
    };
}

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSSortOrderSet {
    ($sorts:expr, $categories:expr, $expanded:expr, $($sort_orders:expr),+) => {
        {
            #[repr(C)]
            #[allow(non_snake_case)]
            struct SortOrderSet {
                sorts: u32,
                categories: u32,
                expanded: u32,
                sort_orders: [SSortOrder; $sorts],
            }
            assert!($categories <= $sorts, "cCategories > cSorts");
            assert!($expanded <= $categories, "cExpanded > cCategories");
            std::mem::transmute::<_, &mut $crate::sys::SSortOrderSet>(&mut SortOrderSet {
                sorts: $sorts,
                categories: $categories,
                expanded: $expanded,
                sort_orders: [$($sort_orders,)+],
            })
        }
    };
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Arc;

    #[test]
    fn login() {
        let initialized = Initialize::new(Default::default()).expect("failed to initialize MAPI");
        let _logon = Logon::new(
            Arc::new(initialized),
            Default::default(),
            None,
            None,
            LogonFlags {
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
