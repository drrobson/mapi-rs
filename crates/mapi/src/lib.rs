//! All of the safe wrappers added by this crate, as well as any macros, are exported from the root
//! module of this crate.
//!
//! All of the nested, unsafe types from `outlook_mapi_sys` are re-exported as the `sys` module in
//! this crate.

pub use outlook_mapi_sys::Microsoft::Office::Outlook::MAPI::Win32 as sys;

pub mod mapi_initialize;
pub mod mapi_logon;
pub mod row;
pub mod row_set;

pub use mapi_initialize::*;
pub use mapi_logon::*;
pub use row::*;
pub use row_set::*;

/// Declare a variable length struct with the same layout as [`sys::ENTRYID`] and cast that to
/// `&mut sys::ENTRYID` for use in APIs that expect `*mut sys::ENTRYID`.
///
/// ```
/// use outlook_mapi::SizedENTRYID;
///
/// let entry_id = unsafe {
///     SizedENTRYID!({
///         flags: [0x0, 0x1, 0x2, 0x3],
///         bytes: [0x4, 0x5, 0x6, 0x7, 0x8, 0x9, 0xa, 0xb, 0xc, 0xd, 0xe, 0xf],
///     })
/// };
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedENTRYID {
    ({ flags: $flags:expr, bytes: $bytes:expr $(,)? }) => {
        std::mem::transmute::<_, &mut $crate::sys::ENTRYID>(&mut {
            #[repr(C)]
            struct EntryId {
                flags: [u8; 4],
                bytes: [u8; $bytes.len()],
            }
            EntryId {
                flags: $flags,
                bytes: $bytes,
            }
        })
    };
}

/// Declare a variable length struct with the same layout as [`sys::SPropTagArray`] and cast that
/// to `&mut sys::SPropTagArray` for use in APIs that expect `*mut sys::SPropTagArray`.
///
/// ```
/// use outlook_mapi::{sys, SizedSPropTagArray};
///
/// let prop_tag_array = unsafe {
///     SizedSPropTagArray!([
///         sys::PR_ENTRYID,
///         sys::PR_DISPLAY_NAME_W,
///     ])
/// };
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSPropTagArray {
    ($tags:expr) => {
        std::mem::transmute::<_, &mut $crate::sys::SPropTagArray>(&mut {
            const COUNT: usize = $tags.len();
            #[repr(C)]
            struct PropTagArray {
                count: u32,
                tags: [u32; COUNT],
            }
            PropTagArray {
                count: COUNT as u32,
                tags: $tags,
            }
        })
    };
}

/// Declare a variable length struct with the same layout as [`sys::SPropProblemArray`] and cast
/// that to `&mut sys::SPropProblemArray` for use in APIs that expect `*mut
/// sys::SPropProblemArray`.
///
/// ```
/// use outlook_mapi::{sys, SizedSPropProblemArray};
///
/// let prop_problem_array = unsafe {
///     SizedSPropProblemArray!([
///         sys::SPropProblem {
///             ulIndex: 0,
///             ulPropTag: sys::PR_ENTRYID,
///             scode: sys::MAPI_E_NOT_FOUND.0
///         },
///         sys::SPropProblem {
///             ulIndex: 1,
///             ulPropTag: sys::PR_DISPLAY_NAME_W,
///             scode: sys::MAPI_E_NOT_FOUND.0
///         },
///     ])
/// };
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSPropProblemArray {
    ($problems:expr) => {
        std::mem::transmute::<_, &mut $crate::sys::SPropProblemArray>(&mut {
            const COUNT: usize = $problems.len();
            #[repr(C)]
            struct ProblemArray {
                count: u32,
                problems: [$crate::sys::SPropProblem; COUNT],
            }
            ProblemArray {
                count: COUNT as u32,
                problems: $problems,
            }
        })
    };
}

/// Declare a variable length struct with the same layout as [`sys::ADRLIST`] and cast that to
/// `&mut sys::ADRLIST` for use in APIs that expect `*mut sys::ADRLIST`.
///
/// ```
/// use std::ptr;
/// use outlook_mapi::{sys, SizedADRLIST};
///
/// let adr_list = unsafe {
///     SizedADRLIST!([
///         sys::ADRENTRY {
///             ulReserved1: 0,
///             cValues: 0,
///             rgPropVals: ptr::null_mut(),
///         },
///         sys::ADRENTRY {
///             ulReserved1: 0,
///             cValues: 0,
///             rgPropVals: ptr::null_mut(),
///         },
///     ])
/// };
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedADRLIST {
    ($entries:expr) => {
        std::mem::transmute::<_, &mut $crate::sys::ADRLIST>(&mut {
            const COUNT: usize = $entries.len();
            #[repr(C)]
            struct AdrList {
                count: u32,
                entries: [$crate::sys::ADRENTRY; COUNT],
            }
            AdrList {
                count: COUNT as u32,
                entries: $entries,
            }
        })
    };
}

/// Declare a variable length struct with the same layout as [`sys::SRowSet`] and cast that to
/// `&mut sys::SRowSet` for use in APIs that expect `*mut sys::SRowSet`.
///
/// ```
/// use std::ptr;
/// use outlook_mapi::{sys, SizedSRowSet};
///
/// let row_set = unsafe {
///     SizedSRowSet!([
///         sys::SRow {
///             ulAdrEntryPad: 0,
///             cValues: 0,
///             lpProps: ptr::null_mut(),
///         },
///         sys::SRow {
///             ulAdrEntryPad: 0,
///             cValues: 0,
///             lpProps: ptr::null_mut(),
///         },
///     ])
/// };
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSRowSet {
    ($rows:expr) => {
        std::mem::transmute::<_, &mut $crate::sys::SRowSet>(&mut {
            const COUNT: usize = $rows.len();
            #[repr(C)]
            struct RowSet {
                count: u32,
                rows: [$crate::sys::SRow; COUNT],
            }
            RowSet {
                count: COUNT as u32,
                rows: $rows,
            }
        })
    };
}

/// Declare a variable length struct with the same layout as [`sys::SSortOrderSet`] and cast that
/// to `&mut sys::SSortOrderSet` for use in APIs that expect `*mut sys::SSortOrderSet`.
///
/// ```
/// use outlook_mapi::{sys, SizedSSortOrderSet};
///
/// let sort_order_set = unsafe {
///     SizedSSortOrderSet!({
///         categories: 1,
///         expanded: 1,
///         sorts: [
///             sys::SSortOrder {
///                 ulPropTag: sys::PR_CONVERSATION_TOPIC_W,
///                 ulOrder: sys::TABLE_SORT_DESCEND,
///             },
///             sys::SSortOrder {
///                 ulPropTag: sys::PR_MESSAGE_DELIVERY_TIME,
///                 ulOrder: sys::TABLE_SORT_CATEG_MAX,
///             },
///             sys::SSortOrder {
///                 ulPropTag: sys::PR_CONVERSATION_INDEX,
///                 ulOrder: sys::TABLE_SORT_ASCEND,
///             },
///         ],
///     })
/// };
/// ```
#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSSortOrderSet {
    ({ categories: $categories:expr, expanded: $expanded:expr, sorts: $sorts:expr $(,)? }) => {
        std::mem::transmute::<_, &mut $crate::sys::SSortOrderSet>(&mut {
            const COUNT: usize = $sorts.len();

            let count = COUNT;
            let categories = $categories;
            let expanded = $expanded;

            assert!(categories <= count);
            assert!(expanded <= categories);

            #[repr(C)]
            struct SortOrderSet {
                count: u32,
                categories: u32,
                expanded: u32,
                sorts: [$crate::sys::SSortOrder; COUNT],
            }
            SortOrderSet {
                count: count as u32,
                categories: categories as u32,
                expanded: expanded as u32,
                sorts: $sorts,
            }
        })
    };
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::{mem, ptr};

    #[test]
    fn sized_macros() {
        assert_eq!(
            mem::size_of::<sys::ENTRYID>(),
            mem::size_of_val(unsafe { SizedENTRYID!({ flags: [0; 4], bytes: [0; 1] }) })
        );
        assert_eq!(
            mem::size_of::<sys::SPropTagArray>(),
            mem::size_of_val(unsafe { SizedSPropTagArray!([sys::PR_NULL]) })
        );
        assert_eq!(
            mem::size_of::<sys::SPropProblemArray>(),
            mem::size_of_val(unsafe {
                SizedSPropProblemArray!([sys::SPropProblem {
                    ulIndex: 0,
                    ulPropTag: sys::PR_NULL,
                    scode: sys::MAPI_E_NOT_FOUND.0
                }])
            })
        );
        assert_eq!(
            mem::size_of::<sys::ADRLIST>(),
            mem::size_of_val(unsafe {
                SizedADRLIST!([sys::ADRENTRY {
                    ulReserved1: 0,
                    cValues: 0,
                    rgPropVals: ptr::null_mut(),
                }])
            })
        );
        assert_eq!(
            mem::size_of::<sys::SRowSet>(),
            mem::size_of_val(unsafe {
                SizedSRowSet!([sys::SRow {
                    ulAdrEntryPad: 0,
                    cValues: 0,
                    lpProps: ptr::null_mut(),
                }])
            })
        );
        assert_eq!(
            mem::size_of::<sys::SSortOrderSet>(),
            mem::size_of_val(unsafe {
                SizedSSortOrderSet!({
                    categories: 0,
                    expanded: 0,
                    sorts: [sys::SSortOrder {
                        ulPropTag: sys::PR_NULL,
                        ulOrder: sys::TABLE_SORT_ASCEND,
                    }],
                })
            })
        );
    }
}
