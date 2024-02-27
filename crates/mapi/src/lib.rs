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
macro_rules! SizedENTRYID {
    ($flags:expr, $bytes:expr) => {
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

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSPropTagArray {
    ($tags:expr) => {
        std::mem::transmute::<_, &mut $crate::sys::SPropTagArray>(&mut {
            #[repr(C)]
            struct PropTagArray {
                count: u32,
                tags: [u32; $tags.len()],
            }
            PropTagArray {
                count: $tags.len() as u32,
                tags: $tags,
            }
        })
    };
}

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSPropProblemArray {
    ($problems:expr) => {
        std::mem::transmute::<_, &mut $crate::sys::SPropProblemArray>(&mut {
            #[repr(C)]
            struct ProblemArray {
                count: u32,
                problems: [$crate::sys::SPropProblem; $problems.len()],
            }
            ProblemArray {
                count: $problems.len() as u32,
                problems: $problems,
            }
        })
    };
}

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedADRLIST {
    ($entries:expr) => {
        std::mem::transmute::<_, &mut $crate::sys::ADRLIST>(&mut {
            #[repr(C)]
            struct AdrList {
                count: u32,
                entries: [$crate::sys::ADRENTRY; $entries.len()],
            }
            AdrList {
                count: $entries.len() as u32,
                entries: $entries,
            }
        })
    };
}

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSRowSet {
    ($rows:expr) => {
        std::mem::transmute::<_, &mut $crate::sys::SRowSet>(&mut {
            #[repr(C)]
            struct RowSet {
                count: u32,
                rows: [$crate::sys::SRow; $rows.len()],
            }
            RowSet {
                count: $rows.len() as u32,
                rows: $rows,
            }
        })
    };
}

#[macro_export]
#[allow(non_snake_case)]
macro_rules! SizedSSortOrderSet {
    ($categories:expr, $expanded:expr, $sorts:expr) => {
        #[allow(unused_comparisons)]
        std::mem::transmute::<_, &mut $crate::sys::SSortOrderSet>(&mut {
            #[repr(C)]
            struct SortOrderSet {
                sorts: u32,
                categories: u32,
                expanded: u32,
                sort_orders: [$crate::sys::SSortOrder; $sorts.len()],
            }

            assert!($categories <= $sorts.len(), "cCategories > cSorts");
            assert!($expanded <= $categories, "cExpanded > cCategories");
            SortOrderSet {
                sorts: $sorts.len() as u32,
                categories: $categories,
                expanded: $expanded,
                sort_orders: $sorts,
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
            mem::size_of_val(unsafe { SizedENTRYID!([0; 4], [0; 1]) })
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
                SizedSSortOrderSet!(
                    0,
                    0,
                    [sys::SSortOrder {
                        ulPropTag: sys::PR_NULL,
                        ulOrder: sys::TABLE_SORT_ASCEND,
                    }]
                )
            })
        );
    }
}
